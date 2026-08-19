# pyVBAharness Implementation Guide

How to work on this codebase without breaking the guarantees it exists to
provide. Written for coding agents and engineers making changes; read it
before editing anything under `src/pyvbaharness/`.

Companion documents: [architecture.md](architecture.md) explains what the
design is and why; [troubleshooting.md](troubleshooting.md) explains what
failures mean for users; [AGENTS.md](../AGENTS.md) states the operating
rules. This guide is the how-to.

## 1. Orientation in five minutes

The harness runs VBA inside real desktop Excel, Word, PowerPoint and
Access, and refuses to hang. It is a
three-process system:

```text
your code
  -> ExcelSession        supervisor: no COM at all, owns the watchdogs
       (pipe: JSON lines)
  -> worker process      all COM lives here, single STA thread
       (COM)
  -> EXCEL.EXE           owned, hidden, killed on any hang
```

The split is the whole point: a COM call into Excel can block forever, and a
deadline can only be enforced from outside the blocked apartment. Everything
else follows from that.

Priority order when goals conflict, and it does not bend: **hang resistance,
then accuracy, then performance.**

| Path | Holds |
| --- | --- |
| `session.py` | supervisor, watchdogs, abort path, public API |
| `pool.py` | N sessions in parallel behind a work queue |
| `worker/hosts/base.py` | every COM call that is not app-specific |
| `worker/hosts/{excel,word,powerpoint,access}.py` | one adapter per app |
| `apps.py` | per-app capabilities, measured; no COM, no pywin32 |
| `worker/watcher.py` | dialog scanning (ctypes only, never COM) |
| `worker/__main__.py` | worker command loop, progress tail |
| `codegen.py` | VBA source generation (support, dispatcher, batch) |
| `numbering.py` | source instrumentation (lines, handlers, coverage) |
| `vbasig.py` | VBA declaration parsing |
| `dialog_policy.py`, `oracle.py`, `ranges.py`, `screenshot.py` | pure helpers |

Pure modules exist so behavior can be tested without Excel. Keep new logic
pure whenever it can be.

## 2. Before you start

```powershell
python -m pyvbaharness doctor --live
```

That checks Excel, pywin32, the "Trust access to the VBA project object
model" setting, and the VBE error-trapping mode, then runs a live smoke
test. If it does not pass, fix the environment before writing code:
failures downstream will be misattributed.

Validation commands you will use constantly:

```powershell
python -m pytest tests/unit -q
python -m pytest tests/live -m live -o addopts="" -q
python benchmarks/run_benchmarks.py
python benchmarks/run_pool_benchmarks.py
```

After any live run, confirm cleanliness. This is not optional; a leak is a
defect:

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue
Get-ChildItem "$env:LOCALAPPDATA\pyvbaharness\sessions"
```

## 3. The Office behavior catalog

Every entry below was measured on Office 365 x64, cost real debugging time,
and has a live test guarding it. Treat this as ground truth. Do not "clean
up" code that references these; the workaround is the feature.

Sections 3.1 to 3.12 were measured against Excel and apply to every host
unless they name a worksheet. Sections 3.13 to 3.18 are differences between
hosts, measured 2026-08-18 by running the same operation on all four.

### 3.1 Dispatch attaches instead of creating

`win32com.client.Dispatch("Excel.Application")` calls `GetActiveObject`
first and returns whatever Excel is already running. Since the harness
kills its Excel on a hang, attaching puts a user's open workbooks in the
blast radius. Use `CoCreateInstance`, and keep the PID-snapshot check in
`OfficeHost._prove_owned` that refuses to proceed unless the instance is
provably new. It uses two signals: the main window handle where the app
exposes one, and otherwise a process-list difference that must name
exactly one new process. Creation is serialized machine-wide by the
CREATE mutex so two sessions cannot each see the other in that diff.

### 3.2 Application.Run breaks in-VBA error trapping

An error raised inside a procedure invoked through `Application.Run` does
not unwind into the calling VBA procedure's `On Error` handler, because the
call crosses a COM boundary. Excel shows its runtime-error dialog instead.
The generated dispatcher therefore calls targets **directly**, which is why
`vbasig` has to parse the declaration (Sub versus Function, exact arity)
before the call can be generated.

### 3.3 Erl is per-procedure

Line numbers alone do not give error lines: read in the dispatcher's
handler, `Erl` reports the dispatcher's own unnumbered lines and returns 0.
Only a handler inside the erroring procedure sees the origin line, so
`numbering.instrument_module` injects one per procedure. Those handlers
accumulate frames, which is where VBA stack traces come from.

### 3.4 Modifying the VBProject resets VBA module-level state

**This one bites repeatedly.** Adding or replacing any module clears every
module-level variable in the project. Anything pushed into the support
module (the progress-file path, the coverage arrays) is silently lost the
moment a dispatcher is regenerated.

The pattern to follow, already implemented for both cases: keep the desired
value on the host, clear a `_..._set` flag in `_write_module`, and re-push
lazily just before `Application.Run` (`_ensure_progress_path`,
`_ensure_coverage`). If you add another piece of VBA-side state, wire it the
same way and add a live test that exercises it *after* a module rewrite.

### 3.5 Large Value2 writes wedge Excel after any macro has run

One `Value2` assignment covering roughly 6000+ cells never returns once a
macro has run in that workbook; every Excel window reports
`IsHungAppWindow`. Not a clean cell-count threshold (50x100 took 0.26 s,
100x50 took 5.5 s, 100x64 hung). All writes go through
`ranges.plan_write_chunks` at 2000 cells per chunk, which is also faster.
Reads are unaffected.

### 3.6 All-optional-parameter properties resolve wrongly

Under late-bound dispatch, `Range("A1").Resize(2, 2)` returns `$B$2`, not
`$A$1:$B$2`: pywin32 invokes a property whose parameters are all optional
with no arguments, then the trailing call becomes the returned object's
default `Item`. Same for `Offset`. Build ranges from explicit corner cells
via `Cells(row, column)`.

### 3.7 Quit does not prove Excel exited

An Excel made visible at any point (a compile check does exactly that)
treats itself as user-launched and outlives its last automation client.
`close()` verifies termination and kills if needed.

### 3.8 PrintWindow blocks on a non-pumping window

`PrintWindow` sends a message, so on the wedged Excel you most want to
photograph it never returns. It wedged the timeout path itself during
development. Two defenses, both required: `IsHungAppWindow` selects `BitBlt`
instead, and every capture runs on a daemon thread with a join timeout
(`screenshot.capture_window_safely`). Never call a blocking Win32 API on a
possibly-hung window from a path that must make progress.

### 3.9 A coverage hit cannot prefix a block opener

`PyVbaCovHit 1, 4: If flag Then` turns a block `If` into a single-line
`If`, and the matching `End If` then fails to compile (surfaced as a modal
compile-error dialog). `numbering._BLOCK_OPENERS` excludes `If`, `For`,
`Do`, `While`, `Select`, `With` from hit prefixing; they are still numbered
so `Erl` mapping is unaffected.

### 3.10 Worksheet cells coerce text

A staged batch argument `"5"` becomes the number 5, `"3/4"` becomes a date,
`"=SUM(A1)"` becomes a formula. `codegen.encode_batch_arg` type-prefixes
every value (`s:`, `i:`, `d:`, `b:`, `e:`) so cells always store literal
text and round-trip exactly. If you add a batch argument type, extend both
the encoder and the VBA `DecodeArg`, and add a round-trip live test.

### 3.11 Excel's own prompts are not #32770 dialogs

"Do you want to save your changes?" is a `NUIDialog`, not a classic
`#32770`. Its controls are drawn inside a `NetUIHWND` surface: enumerating
children yields `NetUIHWND`, `NetUICtrlNotifySink`, and `RICHEDIT60W`, all
returning empty text, with no Win32 `Button` anywhere. The dismissal policy
cannot read or click it, and the class-based watcher never saw it, so a run
blocked this way degraded into a bare timeout.

Detection now keys on modality rather than class: a titled `XLMAIN` window
is disabled exactly while something modal owns Excel. Measured 2026-07-25:
zero disabled samples across 15 scans of a CPU-pegged 3 s macro, against 26
of 28 scans while a save prompt was up. The check requires four consecutive
scans (one second) and stands down for two seconds after a successful
dismissal, so a `#32770` the policy can actually handle is never
misreported.

Two traps when working here. Excel runs hidden, so its `XLMAIN` windows are
invisible and any visible-only window enumeration filters away the very
signal being looked for; `_enum_top_level` therefore returns all windows and
callers filter per use. And these dialogs are still never dismissed: without
readable buttons, clicking blind could save a workbook the caller never
wanted saved.

### 3.12 Excel is not a child of its COM client

DCOM launches it, so killing the worker tree never reaps Excel. The harness
records the PID (from `GetWindowThreadProcessId(app.Hwnd)`) plus its start
time, and additionally places Excel in a kill-on-close job object so worker
death of any kind takes Excel with it.

### 3.13 Word cannot pass arguments into a ParamArray

`Application.Run` in Word fails with `DISP_E_PARAMNOTFOUND` (0x80020007)
whenever the target procedure declares `ParamArray`. Excel, PowerPoint and
Access all accept it. Explicit `ByVal ... As Variant` parameters work on all
four, so `codegen.call_module_source` generates one named parameter per
argument (`pyVbaArg0`, `pyVbaArg1`, ...). The dispatcher is already
regenerated whenever the arity changes, so this costs nothing. Do not
"simplify" it back to a ParamArray: Word runs with arguments stop working
and nothing else does.

### 3.14 Access cannot call Application.Run through pywin32

Every `app.Run(...)` on Access through `win32com.client.dynamic` raises
`DISP_E_PARAMNOTOPTIONAL` (scode 0x8002000E), whatever the target. Padding
the 30 optional parameters with `pythoncom.Missing` does not help, and it
fails identically for a zero-argument target, so it is the dispatch wrapper
rather than the callee. The same call through a raw `IDispatch::Invoke`
returns normally. `AccessHost._invoke_run` therefore resolves the `Run`
DISPID once and invokes it directly.

`Application.Eval` also works and was considered as the fallback (it
returned a 60000-character string intact), but it evaluates expressions, so
it cannot call a `Sub` and would need arguments quoted into the expression
text. The raw Invoke keeps one code path for all four hosts.

### 3.15 Run reference formats differ per host

There is no single string that works everywhere:

| Host | Reference | Notes |
| --- | --- | --- |
| Excel | `'Book1.xlsm'!Module.Proc` | document qualification required |
| Word | `Module.Proc` | document-qualified forms fail |
| PowerPoint | `Module.Proc` | also accepts qualified forms |
| Access | `Proc` | any qualified form fails: "cannot find the procedure" |

Each host supplies its own `_run_ref`. The unqualified forms are safe
because a session owns its instance and keeps one document open in it.

### 3.16 PowerPoint is single-instance and cannot be hidden

`Application.Visible = False` raises "Invalid request. Hiding the
application window is not allowed", so PowerPoint runs on screen; the
`app-created` event reports `can_hide: false` rather than claiming a hidden
host. More seriously, a second `CoCreateInstance` returns the process that
is already running: two activations produced one PID, where Excel, Word and
Access each produced two. A PowerPoint session therefore cannot run
alongside another, cannot be pooled, and cannot start while the user has
PowerPoint open. `_prove_owned` refuses in that case, because taking
ownership of a process the harness did not create would put someone else's
presentation inside the blast radius of a timeout kill.

### 3.17 An unsaved Access module raises a modal prompt at close

A module added through the VBE is unsaved, and Access raises a modal
"Save As / Module Name" dialog for each one when the database closes. That
prompt blocked `Quit` outright during development and left Access wedged
with no way in. `AccessHost.close_document` deletes the components the
session injected before calling `CloseCurrentDatabase`, so the prompt never
exists. Prevention is the pattern to reach for first here: the dialog has no
Win32 buttons, so a watcher could report it but never dismiss it.

The removal is scoped to `self._injected`, never the whole project, because
an opened database's own modules belong to the caller.

### 3.18 Access ordering: SetWarnings needs a database

`DoCmd.SetWarnings False` fails with "The command or action 'SetWarnings'
isn't available now" when no database is open, which is where it would
naturally go in `_configure_app`. It is applied in `_open_finished` instead,
once a database exists. Access also has no unsaved document at all, so
`new_document` creates a scratch `.accdb` in a temp directory, and no
"trust access to the VBA project object model" option, so it has no VBOM
preflight.

## 4. Invariants

These are the guarantees the harness sells. Changing one is a contract
change: update the docs, the oracle, and the tests in the same patch.

1. The harness creates its own instance and never attaches to a running
   one, on any host.
2. Every command carries a positive timeout. A breach kills the recorded
   host process, kills the worker, and marks the session dead.
3. No command runs after the owned host is killed.
4. A timeout or blocked modal is infrastructure state, never evidence about
   the VBA under test. Only `passed` and `vba-error` describe the code.
5. VBA errors are captured inside VBA by a directly-called dispatcher.
6. Dialogs offering a real choice are never dismissed by guessing.
7. All COM lives in the worker; supervisor and watcher stay COM-free.
8. Compile checks serialize machine-wide (they drive the visible VBE).
9. One session serves one caller at a time; `SessionPool` enforces this by
   checkout.
10. Documents open read-only by default and close without saving.
11. A prompt that can be prevented is prevented, never dismissed after the
    fact. See section 4.1.

### 4.1 Prefer prevention, then a deterministic signal, then a deadline

Wedge resistance is built in that order, and new work should follow it.

Prevention comes first because Office's own prompts (NUIDialog and the
`bosa_sdm_*` family) draw their controls inside a NetUI surface: there are
no Win32 buttons to enumerate or click, so a watcher can see them and never
answer them. Anything that stops the dialog existing beats anything that
reacts to it. Current examples: alerts and link prompts off before a
document exists, `FeatureInstall = msoFeatureInstallNone` so a missing
component raises an error rather than the Windows Installer dialog,
AutoRecover off, Word's `ConfirmConversions` off, PowerPoint's `Saved = True`
before close, and Access's injected modules deleted before the database
closes.

A deterministic signal comes second, where one exists. Prefer a state the
OS or the object model will tell you about exactly over a sampled guess:

- `wait_for_exit` blocks on the process handle, which the kernel signals at
  the instant the process ends, instead of polling `is_process_alive`.
- The VBE disables its Compile control precisely when a project finishes
  compiling, so a clean compile returns in milliseconds rather than waiting
  out the dialog window.
- The watcher suppresses its generic modal check while a dialog it clicked
  is still a valid window, rather than for a fixed settle interval.

A deadline is the backstop, not the mechanism. It exists because a COM call
into a blocked apartment cannot be interrupted from inside, which is the one
case no signal can fix. Two waits remain irreducibly time-based, and both
are proving a negative: the compile watch window when neither a dialog nor
the control-disabled signal appears, and the four-scan confirmation before
reporting a modal that is only visible as a disabled main window.

pywin32 exposes no `CoRegisterMessageFilter`, so COM-level "callee is busy"
rejections cannot be retried from inside the worker. They are named in
`_REJECTION_HRESULTS` so the message says the host was busy or showing a
dialog rather than printing a bare HRESULT.

`oracle.py` encodes most of these as a trace validator. Sessions validate
their own trace on close and warn, and the unit suite replays synthetic
traces, so a regression in safety behavior fails loudly rather than
silently.

## 5. Recipe: add a new capability end to end

Worked example, mirroring how `reset_sheets` was added.

**Step 1, name the command.** Add `CMD_RESET_SHEETS = "reset_sheets"` to
`protocol.py`. Add an event kind too if the feature reports asynchronously.

**Step 2, implement the COM work** in `worker/hosts/base.py` if every app
can do it, or in the one adapter that can, decorated so
COM failures become diagnosable `HostError`s:

```python
@_wrap_com("reset sheets")
def reset_sheets(self) -> dict[str, Any]:
    self._require_workbook()
    ...
    return {"cleared_sheets": cleared}
```

Rules for this layer: never trust an all-optional-parameter property (3.6);
route bulk writes through `_write_block` (3.5); if you push state into VBA,
follow the re-push pattern (3.4).

**Step 3, dispatch it** in `worker/__main__.py` inside `Worker._dispatch`.
Return a plain JSON-serializable dict.

**Step 4, expose it** on `ExcelSession` with an explicit timeout argument:

```python
def reset_sheets(self) -> dict[str, Any]:
    return self._expect_passed(
        self._command(protocol.CMD_RESET_SHEETS, {}, None))
```

Use `_expect_passed` for infrastructure commands (it raises on anything but
success). Use the `RunResult` path only for things that carry VBA outcomes.

**Step 5, test both layers.** A unit test through the fake worker for the
plumbing, and a live test for the Excel behavior. If the feature can change
what happens during a hang, add a hang test.

**Step 6, document it.** README section if users call it; architecture.md if
it changes the design; this guide if it adds a landmine.

## 6. Changing VBA codegen safely

The generated VBA is the least forgiving part of the codebase, because a
mistake surfaces as a modal compile dialog rather than a Python traceback.

- The VBA source lives in single-quote-tripled f-strings. The code contains
  runs of doubled double-quotes (VBA string escaping); never introduce three
  consecutive apostrophes into the VBA text.
- Use `Str$` for numeric-to-text, never `CStr`: `CStr` uses the user's
  decimal separator and produces invalid JSON in many locales.
- Sources injected via `AddFromString` must not contain `Attribute` header
  lines. `strip_module_header` removes them.
- Never route a managed run through `Application.Run` from inside VBA
  (3.2).
- Wrap arguments in extra parentheses to force ByVal, so a Variant from the
  ParamArray can satisfy a typed parameter.
- After changing generated source, run the unit codegen tests (they assert
  structure and quoting) AND a live test: only Excel can tell you it
  compiles.
- When an error handler must survive a loop, use the `Resume <label>`
  pattern, as `batch_module_source` does; `On Error GoTo` cannot re-arm
  while error state is pending.

Instrumentation (`numbering.py`) has one rule above all: **preserve the
line count of the original source.** Erl maps back by physical line index,
so a coverage hit rides the numbered line as a colon compound statement
rather than taking a line of its own. The transformer is deliberately
conservative: when unsure, leave the line alone, which only costs precision.

## 7. Testing strategy

**Unit tests (`tests/unit`, no Excel).** Everything pure: dialog policy,
oracle traces, signature parsing, codegen shape, chunk planning,
instrumentation. Plus the supervisor state machine, driven by
`tests/unit/fake_worker.py`, which speaks the pipe protocol and simulates
hangs, blocked modals, and crashes by target-name substring (`Hang`,
`Modal`, `Die`, `Slow`). Add behaviors there rather than reaching for Excel.

**Live tests (`tests/live`, real Excel).** Anything that depends on Excel
actually behaving like Excel. Every landmine in section 3 has one. Live
tests must clean up after themselves; the suite asserts no surviving
processes.

**Writing a good live test.** Prove the mechanism, not the coincidence: use
wall-clock evidence for concurrency (two 2 s runs finishing in under 3.4 s),
assert the harness's own trace where relevant, and always give hangs a
bounded timeout. Remember pytest buffers `-q` output, so a hanging live test
looks like silence; use `-v` and watch `Get-Process EXCEL` CPU when
diagnosing.

## 8. Debugging playbook

When something wedges, do not guess. This sequence found every bug in the
catalog:

1. **Locate the block.** Set `PYVBAHARNESS_STACK_DUMP_S=10` and re-run; the
   worker dumps every thread's stack to stderr, which the session tails.
   That is how the large-write wedge was pinned to the `Value2` assignment.
2. **Check whether Excel is spinning or stuck.** `Get-Process EXCEL` with
   CPU time distinguishes a VBA busy loop (CPU climbing) from a blocked COM
   call (CPU flat). `IsHungAppWindow` on its windows tells you if it is
   pumping messages.
3. **Reproduce standalone.** Write a probe script that does the same COM
   calls with no harness, with a watchdog thread that calls `os._exit` after
   N seconds and prints the stage. Harness code has too many moving parts to
   bisect in place.
4. **Bisect the sequence.** The large-write wedge only reproduced *after a
   macro had run*, and only when no smaller write preceded it. Vary one step
   at a time; a passing run proves nothing until you have matched the exact
   order.
5. **Fix, then guard.** Add a live test that fails without the fix, record
   the measurement (numbers, date, Excel build) in a comment at the
   workaround, and add an entry to section 3.

Never conclude from a passing run that a workaround is unnecessary. Several
of these behaviors appear only in specific orders.

## 9. Performance notes

Current measured costs (Excel 365 x64, Python 3.14,
`benchmarks/output/baseline-1.0.0.json`):

| Operation | Cost |
| --- | --- |
| Session startup and teardown | 0.5 s warm |
| Warm run, same target | 0.5 ms |
| Run with arguments | 0.7 ms |
| `run_vba` with identical source (cache hit) | 0.9 ms |
| Retarget (dispatcher regenerated) | 76 ms |
| Batch, 1000 calls | 0.094 ms per call (6.5x) |
| Compile check, clean project | 1.0 s |
| Write 10,000 cells | 63 ms |
| Read 10,000 cells | 9 ms |

Where the speed comes from, so you do not accidentally remove it:

- **Signature cache** (`OfficeHost._resolved`): resolving a target used to
  read module source through VBE COM on every run. Caching it took warm runs
  from 15 ms to 0.6 ms. Invalidated by `_write_module`.
- **Injection cache** (`ExcelSession._injected`): identical source is never
  resent.
- **Dispatcher cache** (`_call_signature`): regenerated only when the
  (module, proc, arity) changes, which is why retargeting costs 97 ms and
  repeating does not.
- **Batching**: one COM round trip for many calls. The win grows with size
  (2.7x at 50 calls, 9.9x at 3000).

If you add a per-run COM call, measure before and after. A single extra
round trip is roughly a 50% regression on the warm path now.

## 10. Definition of done

Before presenting a change:

- [ ] `python -m pytest tests/unit -q` passes.
- [ ] `python -m pytest tests/live -m live -o addopts="" -q` passes when the
      change touches COM, codegen, dialog policy, process control, or
      instrumentation.
- [ ] No Excel processes and no manifests survive a live run.
- [ ] Benchmarks re-run if the change could affect per-run cost, with the
      baseline JSON updated.
- [ ] New Excel behavior recorded in section 3 with its measurement, and
      guarded by a live test that fails without the fix.
- [ ] Invariants in section 4 intact, or deliberately changed with docs,
      oracle, and tests updated together.
- [ ] Plain ASCII, no em dashes or curly quotes, comments explain
      constraints rather than narrate code.
- [ ] The report states what was actually run and observed, including
      skipped checks.
