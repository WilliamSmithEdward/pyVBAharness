"""Shared Office COM host: everything that is not app-specific.

Rules enforced here (see docs/architecture.md):

- always a NEW application instance via CoCreateInstance with
  CLSCTX_LOCAL_SERVER; never GetActiveObject (a user's open Office document
  is out of bounds, because a timeout kills the instance)
- purely late-bound dynamic dispatch: the pywin32 gencache is never touched
  (a corrupt gencache was observed on the reference machine; dynamic
  dispatch sidesteps that entire failure class)
- prompts are made impossible before they can fire, rather than dismissed
  after; every alert switch is set before a document exists
- the host process id is discovered and proven new before the harness
  touches the application at all, so the supervisor can always kill exactly
  this instance and never someone else's

Subclasses supply the app-specific seams: the ProgID and image name, how the
application is configured, where the VBA project lives, how a document is
created/opened/saved/closed, and how ``Application.Run`` is addressed.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import gc
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client.dynamic
from pywintypes import com_error

from ... import apps, codegen, vbasig
from ...lock import CREATE_MUTEX_NAME, SessionLock
from ...process_control import KillOnCloseJob, process_ids_by_image

# Office constants (hardcoded: no gencache, no typelib import).
MSO_AUTOMATION_SECURITY_LOW = 1
VBEXT_CT_STD_MODULE = 1
VBEXT_CT_CLASS_MODULE = 2
MSO_CONTROL_POPUP = 10
VBE_COMPILE_CONTROL_ID = 578  # "Compile <project>" on the VBE Debug menu

# Creating an instance is snapshot -> CoCreateInstance -> diff. Two sessions
# starting at once would each see the other's new process in the diff, so the
# whole window is serialized machine-wide. Startup is 1-3 s, and COM server
# activation largely serializes anyway, so the cost is small next to owning a
# wrong process.
CREATE_LOCK_TIMEOUT_S = 120.0


class HostError(Exception):
    """A COM operation failed; carries a diagnosable message."""

    def __init__(self, message: str, hresult: int | None = None) -> None:
        super().__init__(message)
        self.hresult = hresult


# COM's own "the callee would not take the call" results. Office returns
# these while it is busy or has a modal dialog up, and the bare HRESULT is
# unreadable, so they are named. pywin32 exposes no CoRegisterMessageFilter,
# so these cannot be retried from inside; the watcher reports the dialog and
# the supervisor's deadline bounds the rest.
_REJECTION_HRESULTS = {
    0x80010001: "the host rejected the call (RPC_E_CALL_REJECTED): it is "
                "busy, or a modal dialog is waiting for input",
    0x8001010A: "the host was busy and did not answer in time "
                "(RPC_E_SERVERCALL_RETRYLATER)",
    0x80010005: "the host was in the middle of another call "
                "(RPC_E_SERVERCALL_RETRYLATER, nested call)",
}


def describe_com_error(err: com_error) -> tuple[str, int | None]:
    hresult = getattr(err, "hresult", None)
    parts: list[str] = []
    if hresult is not None:
        parts.append(f"HRESULT 0x{hresult & 0xFFFFFFFF:08X}")
    excepinfo = getattr(err, "excepinfo", None)
    if excepinfo:
        source = excepinfo[1]
        description = excepinfo[2]
        scode = excepinfo[5]
        if description:
            parts.append(str(description).strip())
        if source:
            parts.append(f"source={source}")
        if scode not in (None, 0) and scode != hresult:
            parts.append(f"scode=0x{scode & 0xFFFFFFFF:08X}")
    elif getattr(err, "strerror", None):
        parts.append(str(err.strerror))
    if hresult is not None:
        rejection = _REJECTION_HRESULTS.get(hresult & 0xFFFFFFFF)
        if rejection:
            parts.append(rejection)
    return ("; ".join(parts) or str(err)), hresult


def _wrap_com(stage: str, tolerate_disconnect: bool = False):
    """Decorator: convert com_error into HostError with the stage named.

    ``tolerate_disconnect`` additionally converts the AttributeError that
    win32com's dynamic dispatch raises when it cannot resolve a member on a
    disconnected object. That happens when the host process has already
    exited: GetIDsOfNames fails, and pywin32 reports it as a missing
    attribute rather than a COM error. Only teardown sets this, because
    swallowing AttributeError anywhere else would hide ordinary typos in
    this file.
    """

    def decorate(func):
        def wrapper(self, *args, **kwargs):
            try:
                return func(self, *args, **kwargs)
            except com_error as err:
                message, hresult = describe_com_error(err)
                if "not trusted" in message.lower():
                    message = f"{message} {self.vbom_hint()}"
                raise HostError(f"{stage}: {message}", hresult) from err
            except AttributeError as err:
                if not tolerate_disconnect:
                    raise
                raise HostError(
                    f"{stage}: the {self.info.label} object is no longer "
                    f"reachable ({err}); the process it belonged to has "
                    "already exited.") from err
        return wrapper

    return decorate


class OfficeHost:
    """Base for the per-application COM hosts.

    Subclasses must set the class attributes and implement the document
    lifecycle plus ``_components`` and ``_run_ref``.
    """

    #: The only identity a subclass must declare; everything else
    #: (ProgID, image name, whether it can hide or run more than one
    #: instance) comes from apps.py so there is one place to correct.
    app_key = ""
    # Trust Center path for the "trust access to the VBA project" setting.
    # Access and Publisher expose no such option, so their hint differs.
    vbom_settings_path = ("File > Options > Trust Center > Trust Center "
                          "Settings > Macro Settings")
    module_extensions = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}
    # Only Excel exposes Application.EnableEvents. Reporting False for
    # the others would put a setting in the trace that was never made,
    # so the oracle checks this invariant per app.
    has_enable_events = False

    def __init__(self) -> None:
        self.info = apps.info(self.app_key)
        self.progid = self.info.progid
        self.image_name = self.info.image_name
        self.document_noun = self.info.document_noun
        self.can_hide = self.info.can_hide
        self.app: Any = None
        self.document: Any = None
        self.pid: int = 0
        self.job_active = False
        self.progress_path: str = ""
        self._progress_set = False
        self._cov_dims: tuple[int, int] | None = None
        self._cov_set = False
        self._job: KillOnCloseJob | None = None
        self._support_installed = False
        self._call_signature: tuple[str, str, int] | None = None
        self._batch_signature: frozenset | None = None
        self._resolved: dict[str, tuple[str, str, vbasig.ProcedureSignature]] = {}
        self._compile_control: Any = None
        # Names this session added to the VBA project. Hosts whose
        # teardown must delete them (Access prompts for unsaved
        # modules) need to touch only these, never the caller's own.
        self._injected: set[str] = set()

    def vbom_hint(self) -> str:
        return (f"{self.info.label} refused programmatic access to the VBA "
                f"project. Enable: {self.vbom_settings_path} > tick 'Trust "
                "access to the VBA project object model'.")

    # ----- lifecycle -------------------------------------------------------

    @_wrap_com("create application")
    def create(self) -> dict[str, Any]:
        """Create a brand-new instance and prove it is new before touching it.

        ``win32com.client.Dispatch("Excel.Application")`` must not be used
        here: given a ProgID string it first calls ``pythoncom.connect``,
        which is GetActiveObject, so it silently ATTACHES to whatever is
        already running (observed live, 2026-07-25: two Dispatch calls
        returned the same PID). Attaching would put a user's own documents
        inside the harness's blast radius. CoCreateInstance always launches a
        fresh server, and the ownership proof below turns any regression into
        a refusal instead of a silent hijack.

        Nothing is configured until ownership is proven: setting Visible or
        DisplayAlerts on a stranger's application would already be damage.
        """
        pythoncom.CoInitialize()
        with SessionLock(timeout_s=CREATE_LOCK_TIMEOUT_S,
                         name=CREATE_MUTEX_NAME, purpose="create"):
            before = process_ids_by_image(self.image_name)
            dispatch = pythoncom.CoCreateInstance(
                pywintypes.IID(self.progid), None,
                pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)
            self.app = win32com.client.dynamic.Dispatch(dispatch)
            after = process_ids_by_image(self.image_name)
            self.pid = self._prove_owned(before, after)

        # Tie the application's lifetime to this worker process: if the worker
        # dies for any reason, the kernel kills it (kill-on-close job).
        # Assignment can fail under restrictive job policies; the manifest
        # sweep remains as the fallback for that case.
        self._job = KillOnCloseJob()
        self.job_active = self._job.assign(self.pid)

        self._configure_app()
        version, build = self._version_info()
        return {
            "app": self.app_key,
            "pid": self.pid,
            "attached": False,
            "visible": not self.can_hide,
            "display_alerts": False,
            "enable_events": False if self.has_enable_events else None,
            "job_kill_on_close": self.job_active,
            "can_hide": self.can_hide,
            "app_version": version,
            "app_build": build,
        }

    def _prove_owned(self, before: set[int], after: set[int]) -> int:
        """Return the PID of the instance just created, or refuse.

        Two independent signals, because neither alone is sufficient. The
        window handle authoritatively ties an Application object to a
        process but is not exposed by every app (Word has no
        Application.Hwnd). The process-list difference always works but
        could in principle catch a process the user started at the same
        moment, so it is only trusted when it names exactly one new process.
        """
        new = after - before
        hwnd = self._app_hwnd()
        if hwnd:
            pid = self._pid_from_hwnd(hwnd)
            if pid and pid not in before:
                return pid
            self.app = None
            raise HostError(
                f"Refusing to run: the new {self.info.label} Application "
                f"resolved to already-running process {pid}. The harness "
                "must own its instance because it kills that process on a "
                "hang.")
        if len(new) == 1:
            return new.pop()
        self.app = None
        if not new:
            if not self.info.multi_instance:
                raise HostError(apps.single_instance_reason(self.app_key))
            raise HostError(
                f"Refusing to run: creating a {self.info.label} Application "
                f"started no new {self.image_name} process, so it attached "
                "to one that was already running. The harness must own its "
                "instance because it kills that process on a hang.")
        raise HostError(
            f"Refusing to run: {len(new)} new {self.image_name} processes "
            f"appeared while creating the {self.info.label} Application, so "
            "the one it owns cannot be identified. Close other instances and "
            "retry.")

    @staticmethod
    def _pid_from_hwnd(hwnd: int) -> int:
        pid = wt.DWORD(0)
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value

    def _app_hwnd(self) -> int:
        """Main-window handle, or 0 when the app does not expose one."""
        try:
            return int(self.app.Hwnd)
        except (com_error, AttributeError, TypeError, ValueError):
            return 0

    def _version_info(self) -> tuple[str, str]:
        version = build = ""
        try:
            version = str(self.app.Version)
        except com_error:
            pass
        try:
            build = str(self.app.Build)
        except com_error:
            pass
        return version, build

    def _configure_app(self) -> None:
        """Set every prompt-suppressing switch, before a document exists."""
        raise NotImplementedError

    def _set_quietly(self, name: str, value: Any) -> bool:
        """Best-effort property set; False when the app rejects it."""
        try:
            setattr(self.app, name, value)
            return True
        except com_error:
            return False

    def _resuppress_alerts(self) -> None:
        """Re-assert alert suppression before a harness-initiated operation.

        User VBA can turn alerts back on and leave them that way, which would
        let a later harness Close or SaveAs raise a prompt that nothing can
        dismiss (Office's own prompts carry no Win32 buttons). Cheap
        insurance on the infrequent operations that can prompt; the hot run
        path does not pay for it.
        """
        raise NotImplementedError

    @_wrap_com("close document")
    def close_document(self) -> None:
        raise NotImplementedError

    @_wrap_com("quit application", tolerate_disconnect=True)
    def quit(self) -> None:
        if self.app is not None:
            self._resuppress_alerts()
            self._quit_app()

    def _quit_app(self) -> None:
        self.app.Quit()

    def release(self) -> None:
        """Drop COM references; hangs here are covered by the supervisor's
        cleanup watchdog, not by local timeouts."""
        self.document = None
        self._compile_control = None
        self.app = None
        gc.collect()
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # ----- documents -------------------------------------------------------

    def new_document(self) -> dict[str, Any]:
        raise NotImplementedError

    def open_document(self, path: str, read_only: bool) -> dict[str, Any]:
        raise NotImplementedError

    def save_as(self, path: str) -> dict[str, Any]:
        raise NotImplementedError

    def _document_name(self) -> str:
        return str(self.document.Name)

    def _require_document(self) -> None:
        if self.document is None:
            raise HostError(
                f"No {self.document_noun} is open in this session.")

    # ----- VBA project -----------------------------------------------------

    def _components(self) -> Any:
        """The VBComponents collection holding the harness's modules."""
        raise NotImplementedError

    def _reset_injection_state(self) -> None:
        self._support_installed = False
        self._call_signature = None
        self._batch_signature = None
        self._progress_set = False
        self._resolved = {}
        self._injected = set()

    @_wrap_com("add module")
    def add_module(self, name: str, source: str, kind: str) -> dict[str, Any]:
        codegen.validate_module_name(name)
        return self._write_module(name, source, kind)

    def _write_module(self, name: str, source: str,
                      kind: str) -> dict[str, Any]:
        """Create or replace a module. Skips the reserved-name check so the
        harness can inject its own modules."""
        self._require_document()
        self._resolved = {}  # module change invalidates cached signatures
        # Changing the VBProject resets VBA module-level state, so anything
        # pushed into the support module (progress path, coverage arrays)
        # is gone and must be re-pushed before the next run. Observed live
        # twice: progress events stopped and coverage hits vanished because
        # the dispatcher module was written after those were set.
        self._progress_set = False
        self._cov_set = False
        component_kind = (VBEXT_CT_CLASS_MODULE if kind == "class"
                          else VBEXT_CT_STD_MODULE)
        body = codegen.strip_module_header(source)
        components = self._components()
        existing = self._find_component(components, name)
        if existing is not None:
            components.Remove(existing)
        component = components.Add(component_kind)
        component.Name = name
        if body.strip():
            component.CodeModule.AddFromString(body)
        self._injected.add(name)
        self._after_module_write(name)
        return {"name": name, "kind": kind, "lines": body.count("\n") + 1}

    def _after_module_write(self, name: str) -> None:
        """Hook for apps that must register a module beyond the VBE."""

    @_wrap_com("remove module")
    def remove_module(self, name: str) -> dict[str, Any]:
        self._require_document()
        components = self._components()
        component = self._find_component(components, name)
        if component is None:
            return {"name": name, "removed": False}
        components.Remove(component)
        self._injected.discard(name)
        if name.lower() == codegen.SUPPORT_MODULE_NAME.lower():
            self._support_installed = False
        if name.lower() == codegen.CALL_MODULE_NAME.lower():
            self._call_signature = None
        return {"name": name, "removed": True}

    @staticmethod
    def _find_component(components: Any, name: str) -> Any:
        wanted = name.lower()
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            if str(component.Name).lower() == wanted:
                return component
        return None

    def _module_source(self, component: Any) -> str:
        code_module = component.CodeModule
        count = int(code_module.CountOfLines)
        if count <= 0:
            return ""
        return str(code_module.Lines(1, count))

    def _resolve_target(self, target: str
                        ) -> tuple[str, str, vbasig.ProcedureSignature]:
        """Find the target's module and parse its declaration (cached).

        A bare ``Proc`` target is searched across every component; harness
        modules are skipped so the search cannot resolve into the injected
        dispatcher. The cache is invalidated whenever a module changes.
        """
        cached = self._resolved.get(target.lower())
        if cached is not None:
            return cached
        resolved = self._resolve_target_uncached(target)
        self._resolved[target.lower()] = resolved
        return resolved

    def _resolve_target_uncached(self, target: str
                                 ) -> tuple[str, str,
                                            vbasig.ProcedureSignature]:
        parts = target.split(".")
        components = self._components()
        reserved = {n.lower() for n in codegen.HARNESS_MODULE_NAMES}
        if len(parts) == 2:
            module_name, proc_name = parts
            component = self._find_component(components, module_name)
            if component is None:
                raise HostError(
                    f"Module {module_name!r} does not exist in this "
                    f"{self.document_noun}'s VBA project.")
            signature = vbasig.find_procedure(self._module_source(component),
                                              proc_name)
            if signature is None:
                raise HostError(
                    f"{module_name}.{proc_name} is not a callable Sub, "
                    "Function, or Property Get in that module.")
            return str(component.Name), proc_name, signature

        proc_name = parts[0]
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            name = str(component.Name)
            if name.lower() in reserved:
                continue
            signature = vbasig.find_procedure(self._module_source(component),
                                              proc_name)
            if signature is not None:
                return name, proc_name, signature
        raise HostError(
            f"No module in this {self.document_noun} declares a callable "
            f"{proc_name!r}.")

    def _ensure_support(self) -> None:
        if self._support_installed:
            return
        self._write_module(codegen.SUPPORT_MODULE_NAME,
                           codegen.support_module_source(), "standard")
        self._support_installed = True

    def _ensure_progress_path(self) -> None:
        """Push the progress-file path into VBA, after all module writes.

        The Office process is not a child of the worker, so environment
        variables cannot carry the path. It must be (re)pushed whenever the
        VBProject changed, because that resets VBA module-level state.
        """
        if not self.progress_path or self._progress_set:
            return
        self.run_support("PyVbaSetProgressPath", [self.progress_path])
        self._progress_set = True

    def _ensure_coverage(self) -> None:
        """Re-arm the coverage arrays after a VBProject change.

        Same landmine as the progress path: writing a module resets VBA
        module-level state, which clears mCovReady and the hit arrays, so
        every hit after a dispatcher rewrite would be dropped silently.
        Re-initializing also zeroes accumulated hits, which is honest: VBA
        already discarded them.
        """
        if self._cov_dims is None or self._cov_set:
            return
        modules, max_line = self._cov_dims
        self.run_support("PyVbaCovInit", [modules, max_line])
        self._cov_set = True

    # ----- running ---------------------------------------------------------

    def _run_ref(self, module: str, proc: str) -> str:
        """The string ``Application.Run`` wants for a harness-owned proc."""
        raise NotImplementedError

    def _invoke_run(self, ref: str, args: list[Any]) -> Any:
        return self.app.Run(ref, *args)

    @_wrap_com("call support procedure")
    def run_support(self, proc: str, args: list[Any]) -> Any:
        """Internal Run into the support module (coverage, progress path)."""
        self._require_document()
        return self._invoke_run(
            self._run_ref(codegen.SUPPORT_MODULE_NAME, proc), args)

    @_wrap_com("install support module")
    def ensure_support_module(self) -> dict[str, Any]:
        """Install the harness support module without running anything.

        Compile checks of code written for the harness need PyVbaLog and the
        assert helpers to resolve; running code gets them injected
        automatically, a bare compile does not.
        """
        self._require_document()
        self._ensure_support()
        return {"name": codegen.SUPPORT_MODULE_NAME}

    def _ensure_dispatcher(self, module: str, proc: str,
                           signature: vbasig.ProcedureSignature,
                           arg_count: int) -> None:
        key = (module.lower(), proc.lower(), arg_count)
        if self._call_signature == key:
            return
        self._write_module(
            codegen.CALL_MODULE_NAME,
            codegen.call_module_source(module, proc, signature, arg_count),
            "standard")
        self._call_signature = key

    @_wrap_com("run VBA")
    def run(self, target: str, args: list[Any]) -> str:
        """Managed run through the generated dispatcher.

        Returns the dispatcher's JSON string. The dispatcher calls the target
        directly (not through Application.Run) so a VBA error unwinds into
        its On Error handler instead of raising the host's runtime dialog.
        """
        self._require_document()
        codegen.validate_run_target(target)
        if len(args) > codegen.MAX_RUN_ARGS:
            raise HostError(
                f"run supports at most {codegen.MAX_RUN_ARGS} arguments, "
                f"got {len(args)}.")
        module, proc, signature = self._resolve_target(target)
        if not signature.accepts(len(args)):
            raise HostError(
                f"{module}.{proc} takes {signature.arity_text()} "
                f"argument(s); {len(args)} were supplied.")
        self._ensure_support()
        self._ensure_dispatcher(module, proc, signature, len(args))
        self._ensure_progress_path()
        self._ensure_coverage()
        return str(self._invoke_run(
            self._run_ref(codegen.CALL_MODULE_NAME, codegen.RUNNER_ENTRY),
            args))

    @_wrap_com("run VBA (raw)")
    def run_raw(self, target: str, args: list[Any]) -> Any:
        """Direct Application.Run without the runner (no in-VBA trapping)."""
        self._require_document()
        codegen.validate_run_target(target)
        parts = target.split(".")
        if len(parts) == 2:
            ref = self._run_ref(parts[0], parts[1])
        else:
            ref = self._bare_run_ref(parts[0])
        return self._invoke_run(ref, args)

    def _bare_run_ref(self, proc: str) -> str:
        return proc

    def run_batch(self, calls: list[dict[str, Any]]) -> str:
        raise HostError(
            f"Batch execution is not available for {self.app_key}: it stages "
            "arguments on a hidden worksheet, which only Excel has. Use "
            "run() per call.")

    # ----- module export and coverage --------------------------------------

    @_wrap_com("export modules")
    def export_modules(self, directory: str) -> dict[str, Any]:
        """Export every non-harness component via VBIDE Export."""
        self._require_document()
        reserved = {n.lower() for n in codegen.HARNESS_MODULE_NAMES}
        exported: list[str] = []
        components = self._components()
        for index in range(1, int(components.Count) + 1):
            component = components.Item(index)
            name = str(component.Name)
            if name.lower() in reserved:
                continue
            extension = self.module_extensions.get(int(component.Type))
            if extension is None:
                continue
            target = str(Path(directory) / f"{name}{extension}")
            component.Export(target)
            exported.append(target)
        return {"files": exported}

    def coverage_init(self, module_count: int, max_line: int) -> None:
        self._ensure_support()
        self._cov_dims = (module_count, max_line)
        self._cov_set = False
        self._ensure_coverage()

    def coverage_report(self) -> str:
        self._ensure_support()
        return str(self.run_support("PyVbaCovReportJson", []))

    # ----- visibility and compile ------------------------------------------

    @_wrap_com("set visibility")
    def set_visible(self, visible: bool) -> dict[str, Any]:
        if not visible and not self.can_hide:
            # Reporting success for something that did not happen would make
            # the trace lie; the supervisor surfaces this as a no-op.
            return {"visible": True, "supported": False}
        self.app.Visible = visible
        return {"visible": bool(self.app.Visible), "supported": True}

    @_wrap_com("list procedures")
    def list_procs(self, module: str) -> list[dict[str, Any]]:
        self._require_document()
        component = self._find_component(self._components(), module)
        if component is None:
            raise HostError(
                f"Module {module!r} does not exist in this "
                f"{self.document_noun}'s VBA project.")
        found = []
        for signature in vbasig.list_procedures(
                self._module_source(component)):
            found.append({
                "name": signature.name,
                "kind": signature.kind,
                "required": signature.required,
                "optional": signature.optional,
                "param_array": signature.has_param_array,
            })
        return found

    @_wrap_com("compile project")
    def start_compile(self) -> str:
        """Make the VBE visible and execute its Compile command control.

        The host must be visible for the Compile error dialog to surface as a
        detectable window (XLIDE oracle lesson: hidden hosts convert compile
        rejections into silent false accepts). Returns "already-compiled"
        when the Compile control is disabled, which the VBE does exactly when
        the project is fully compiled; otherwise fires the compile and
        returns "fired". The caller owns the watch window and the verdict.
        """
        self._require_document()
        self._set_quietly("Visible", True)
        vbe = self.app.VBE
        vbe.MainWindow.Visible = True
        self._activate_vbproject(vbe)
        control = self._find_compile_control(vbe)
        if control is None:
            raise HostError(
                "Could not locate the VBE Compile command (control id 578).")
        if not control.Enabled:
            return "already-compiled"
        control.Execute()
        return "fired"

    def _activate_vbproject(self, vbe: Any) -> None:
        """Point the VBE at the project the harness owns."""
        raise NotImplementedError

    def compile_control_disabled(self) -> bool:
        """True when the Compile control has gone disabled.

        The VBE disables its Compile command precisely when the project is
        fully compiled, so after a fired compile this is a POSITIVE success
        signal: the accept verdict no longer has to wait out the full
        dialog-watch window (which existed only to prove a negative).
        """
        try:
            control = self._compile_control
            return control is not None and not control.Enabled
        except com_error:
            return False

    def _find_compile_control(self, vbe: Any) -> Any:
        if self._compile_control is not None:
            return self._compile_control
        bars = vbe.CommandBars
        for bar_index in range(1, int(bars.Count) + 1):
            found = self._search_controls(bars.Item(bar_index).Controls, 0)
            if found is not None:
                self._compile_control = found
                return found
        return None

    def _search_controls(self, controls: Any, depth: int) -> Any:
        if depth > 3:
            return None
        for index in range(1, int(controls.Count) + 1):
            control = controls.Item(index)
            try:
                if int(control.Id) == VBE_COMPILE_CONTROL_ID:
                    return control
                if int(control.Type) == MSO_CONTROL_POPUP:
                    found = self._search_controls(control.Controls, depth + 1)
                    if found is not None:
                        return found
            except com_error:
                continue
        return None

    @_wrap_com("hide VBE")
    def end_compile(self) -> None:
        """Put the host back to hidden, main window before the VBE.

        Order matters. Hiding the VBE first and the host second makes the
        VBE window reappear on its own: measured 2026-08-18, hidden at
        t=1.02 s and visible again at t=1.17 s, after which the watcher
        reported a debugger break and killed a healthy session on a later
        run. Hiding the host first removes what pulls the VBE back up.
        """
        if self.can_hide:
            self._set_quietly("Visible", False)
        self.hide_vbe()

    def hide_vbe(self) -> bool:
        """Hide the VBE main window; True when it is confirmed hidden."""
        try:
            window = self.app.VBE.MainWindow
            window.Visible = False
            return not bool(window.Visible)
        except com_error:
            return False
