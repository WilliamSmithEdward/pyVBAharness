# Troubleshooting

Start with the built-in diagnosis; it checks everything on this page that a
registry or module can reveal:

```bash
python -m pyvbaharness doctor --live
```

## "Programmatic access to the Visual Basic Project is not trusted"

The harness injects VBA modules through the VBA project object model, which
Excel blocks by default.

Turn it on: **File > Options > Trust Center > Trust Center Settings > Macro
Settings > Trust access to the VBA project object model**.

The setting is per Office application and per Windows user, so enabling it in
Excel does not affect Word, and enabling it for you does not affect other
accounts. If the harness runs under a different account (a service, a
scheduled task, a CI agent), enable it under that account.

## The first run times out, later runs are fine

Excel's first COM launch after a reboot, an update, or an activation prompt
can take longer than the startup deadline. Treat one cold timeout as warmup:
retry, and expect subsequent runs to be quick. A systematic timeout after
warmup is a real problem, not cold start.

To give a slow machine more room:

```python
from pyvbaharness import ExcelSession, HarnessConfig

excel = ExcelSession(HarnessConfig(startup_timeout_s=120))
```

## "Another pyvbaharness session is running on this machine"

Sessions take a machine-wide lock because Excel's automation surface is
shared: two sessions contend over modal dialogs, VBE focus, and process
cleanup, producing timeouts that look like code failures. Wait for the other
session to finish.

If the lock is held by a session that crashed, the mutex is released by
Windows automatically when that process dies. If the message persists, look
for a stray `python -m pyvbaharness.worker` process.

To wait rather than fail immediately:

```python
ExcelSession(HarnessConfig(lock_wait_s=300))
```

## Every run reports `modal-blocked` at the first VBA error

Check the VBE's error-trapping mode (`doctor` reports it). Tools > Options >
General > Error Trapping in the VBE: "Break on All Errors" stops in the
debugger even for errors the harness traps, so the VBE window takes over and
the run is aborted as blocked. Set "Break on Unhandled Errors".

## `result.error.line` is None or points at an earlier line

Line capture needs the instrumentation that `run_vba` applies by default.
It is None when `line_numbers=False` was passed, when the source already
contained numeric line labels (mixing would be a compile error, so the
harness declines), or when the error fired before any numbered line ran.
A line the numberer conservatively skips (a structure keyword such as
`Case`, or a `Dim`) attributes to the nearest numbered line above it.

## A run reports `modal-blocked`

Excel showed a dialog that needs a human decision (Yes/No, Save changes?,
Retry/Cancel, or a compile error). The harness does not guess at those, so it
killed the Excel instance and reported the dialog in `result.dialogs`.

Common causes:

- VBA calling `MsgBox` with buttons other than OK. Use `PyVbaLog` instead;
  its output comes back in `result.output`.
- Code that saves, closes, or opens workbooks and hits a confirmation.
- A compile error in the project. Run `compile_project()` to see the text.

## A run reports `runner-error`

The harness or COM failed around the run rather than inside VBA. The message
names the stage. Frequent cases:

- `Module 'X' does not exist in this workbook's VBA project`: the target
  module was never added, or the workbook was replaced by a later
  `new_workbook()` or `open_workbook()` call, which clears injected modules.
- `X.Y is not a callable Sub, Function, or Property Get`: the target is a
  `Property Let` or `Property Set`, or the name is misspelled.
- `X.Y takes 2 argument(s); 3 were supplied`: arity is checked against the
  parsed declaration before Excel is touched, so this never becomes a VBA
  compile error.

## A run reports `timeout`

The run passed its deadline and the Excel process was killed. A timeout says
nothing about whether the VBA is correct; it says the run did not finish.

Check first whether the code genuinely blocks: an infinite loop, a modal the
watcher could not see, a `Sleep`, or a network call inside VBA. If the work
is simply slow, raise the deadline for that call:

```python
excel.run_macro("Model.Recalculate", timeout=600)
```

## Excel processes pile up

They should not, even across crashes. Each owned Excel sits in a kernel job
object with kill-on-close, so the kernel terminates it the moment its worker
process dies, cooperatively or not. On top of that, each session records its
Excel process id and start time, kills that process on any abnormal outcome,
verifies termination on close, and sweeps orphans recorded by crashed
sessions at the next start.

To check, look for hidden Excel processes:

```powershell
Get-Process EXCEL | Select-Object Id, StartTime, MainWindowTitle
```

An Excel started by the harness has no main window title. To make the sweep
run immediately rather than on the next session, delete the leftover
manifests in `%LOCALAPPDATA%\pyvbaharness\sessions` after killing the
processes.

## A returned value is `"<object:Range>"` instead of data

VBA returned an object. Objects cannot cross the JSON boundary, so they are
reported as a marker rather than failing the run. Return a scalar or an
array, or have the VBA write into cells and read them with `read_range`.

## Diagnosing a wedged run

Set `PYVBAHARNESS_STACK_DUMP_S` to a number of seconds. The worker then dumps
every thread's stack to stderr at that interval, which the session captures.
This is how the large-write wedge documented in the architecture notes was
located.

```powershell
$env:PYVBAHARNESS_STACK_DUMP_S = "10"
```

## Known Excel behaviors the harness works around

These were measured on Excel 365 x64 and are handled for you; they are listed
so surprising behavior in your own automation code is recognizable.

- A single `Value2` assignment covering roughly 6000 or more cells wedges
  Excel indefinitely once any macro has run in that workbook. Writes are
  split into 2000-cell blocks, which is also faster.
- `Range("A1").Resize(2, 2)` returns cell `B2` under late-bound Python COM,
  not `A1:B2`, because a property whose parameters are all optional is
  invoked with no arguments and the trailing call becomes the default item
  indexer. Build ranges from explicit corner cells instead.
- An error raised inside a procedure invoked through `Application.Run` does
  not reach the calling VBA procedure's error handler, so Excel shows its
  runtime-error dialog. The harness calls targets directly instead.
- `Dispatch("Excel.Application")` attaches to a running Excel rather than
  starting a new one. The harness uses `CoCreateInstance` and verifies the
  resulting process did not already exist.
- `Application.Quit()` can return successfully without the process ending.
  An Excel that was made visible at any point (a compile check does that)
  treats itself as user-launched and outlives its last automation client.
  Closing a session verifies the process actually ended and kills it if not.
