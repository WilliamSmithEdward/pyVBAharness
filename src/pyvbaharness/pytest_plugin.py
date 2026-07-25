"""pytest plugin: VBA test procedures as first-class pytest items.

Files named ``test_*.bas`` (or ``*_test.bas``) are collected; each
zero-argument ``Test*`` Sub/Function becomes a pytest item that runs through
one shared, auto-recycling ExcelSession. VBA assertion failures
(PyVbaAssert / PyVbaAssertEqual) report their message, error line, stack,
and PyVbaLog output; a hanging test reports timeout and costs one recycle,
not the session.

pytest gives the rest for free: ``-k`` selection, ``-m vba``, JUnit XML for
CI, and per-worker isolation under pytest-xdist (each worker process gets
its own non-exclusive session).

ini options:
  vba_collect       enable .bas collection (default: true)
  vba_test_timeout  per-test timeout in seconds (default: 30)
  vba_worker_argv   test seam: alternate worker argv (spaces separated)
"""
from __future__ import annotations

from pathlib import Path

import pytest

from . import vbasig

_SESSION_KEY = pytest.StashKey["object"]()


def pytest_addoption(parser) -> None:
    parser.addini("vba_collect", type="bool", default=True,
                  help="collect test_*.bas files as VBA test suites")
    parser.addini("vba_test_timeout", default="30",
                  help="per-VBA-test timeout in seconds")
    parser.addini("vba_worker_argv", default="",
                  help="override the harness worker argv (testing seam)")


def pytest_configure(config) -> None:
    config.addinivalue_line(
        "markers", "vba: VBA test procedure run through pyvbaharness")


def pytest_collect_file(file_path: Path, parent):
    if not parent.config.getini("vba_collect"):
        return None
    name = file_path.name.lower()
    if not name.endswith(".bas"):
        return None
    if not (name.startswith("test_") or name.endswith("_test.bas")):
        return None
    return VbaFile.from_parent(parent, path=file_path)


def _get_session(config):
    session = config.stash.get(_SESSION_KEY, None)
    if session is not None:
        return session
    from .session import ExcelSession, HarnessConfig

    argv_text = str(config.getini("vba_worker_argv") or "").strip()
    harness_config = HarnessConfig(
        exclusive=False,  # xdist workers must coexist
        auto_recycle=True,
        worker_argv=argv_text.split() if argv_text else None,
    )
    session = ExcelSession(harness_config)
    config.stash[_SESSION_KEY] = session
    config.add_cleanup(session.close)
    return session


class VbaFile(pytest.File):
    def collect(self):
        source = self.path.read_text(encoding="utf-8-sig", errors="replace")
        module = self.path.stem
        from .codegen import is_vba_identifier

        if not is_vba_identifier(module):
            return
        for test_name in vbasig.discover_tests(source, prefix="Test"):
            yield VbaTestItem.from_parent(self, name=test_name,
                                          module=module, source=source)


class VbaTestFailure(Exception):
    def __init__(self, result) -> None:
        super().__init__(str(result.error) if result.error
                         else result.message)
        self.result = result


class VbaInfrastructureFailure(Exception):
    pass


class VbaTestItem(pytest.Item):
    def __init__(self, *, module: str, source: str, **kwargs) -> None:
        super().__init__(**kwargs)
        self.module = module
        self.source = source
        self.add_marker(pytest.mark.vba)

    def runtest(self) -> None:
        session = _get_session(self.config)
        timeout = float(self.config.getini("vba_test_timeout"))
        if not session.has_workbook:
            session.new_workbook()
        # Cache no-op normally; reinjection after a prior test's recycle.
        session.add_module(self.module, self.source, line_numbers=True)
        result = session.run_macro(f"{self.module}.{self.name}",
                                   timeout=timeout)
        if result.outcome == "passed":
            return
        if result.outcome == "vba-error":
            raise VbaTestFailure(result)
        raise VbaInfrastructureFailure(
            f"{result.outcome}: {result.message or 'the run did not finish'}")

    def repr_failure(self, excinfo):
        if isinstance(excinfo.value, VbaTestFailure):
            result = excinfo.value.result
            lines = [str(result.error)]
            if result.error is not None and result.error.stack:
                lines.append("VBA stack (deepest first):")
                lines.extend(f"  {proc} line {line}"
                             for proc, line in result.error.stack)
            if result.output:
                lines.append("PyVbaLog output:")
                lines.extend(f"  {entry}" for entry in result.output)
            return "\n".join(lines)
        if isinstance(excinfo.value, VbaInfrastructureFailure):
            return (f"harness infrastructure failure: {excinfo.value} "
                    "(not a verdict on the VBA under test)")
        return super().repr_failure(excinfo)

    def reportinfo(self):
        return self.path, 0, f"[vba] {self.module}.{self.name}"
