# pyVBAharness

Run VBA in desktop Excel, Word, PowerPoint and Access from Python, under a
supervisor that enforces a deadline on every call.

Office automation has three failure modes that ordinary error handling does
not cover:

- A VBA runtime error opens a modal dialog and waits indefinitely.
- `Application.Run` takes no timeout parameter.
- An Office application created over COM is not a child process, so
  terminating the caller leaves `EXCEL.EXE` running with the document open.

The harness addresses each one. VBA errors are trapped inside VBA and
returned as data. Every command is issued from a supervisor process that
holds no COM references, so it can always enforce a deadline. When a
deadline expires, the process the harness owns is terminated by recorded
process ID.

```python
from pyvbaharness import ExcelSession

with ExcelSession() as excel:
    result = excel.run_vba("""
Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function
""", proc="AddNums", args=(20, 22))

    result.outcome   # 'passed'
    result.value     # 42
```

The same API drives the other three hosts:

```python
from pyvbaharness import AccessSession, PowerPointSession, WordSession

with WordSession() as word:
    word.run_vba(source, proc="Main")
```

## Supported applications

| | Excel | Word | PowerPoint | Access |
| --- | --- | --- | --- | --- |
| Class | `ExcelSession` | `WordSession` | `PowerPointSession` | `AccessSession` |
| Runs hidden | yes | yes | no | yes |
| Concurrent sessions and `SessionPool` | yes | yes | no | yes |
| Needs the Trust Center setting below | yes | yes | yes | no |
| `save_as` | `.xlsm`, `.xlsb` | `.docm`, `.dotm` | `.pptm`, `.potm` | not applicable |
| Ranges and `run_batch` | yes | no | no | no |

Injection, managed runs, trapped VBA errors with source lines and stacks,
coverage, progress reporting, timeouts, compile checks, module export and
the pytest plugin work identically on all four.

Two PowerPoint limits are properties of PowerPoint, not choices. It refuses
`Application.Visible = False`, so its runs happen on screen. A second COM
activation returns the process that is already running rather than starting
a new one, so a PowerPoint session cannot run alongside another one or
inside a `SessionPool`, and cannot start at all while you have PowerPoint
open. The harness refuses in that case instead of taking ownership of a
process it did not create, because a timeout kills the process it owns.

Access differs in two ways worth knowing before use. It has no unsaved
document, so `new_document()` creates a scratch `.accdb` in a temp directory
and deletes it at teardown. Injecting a module writes into the database file
immediately rather than at save time, so `open_document` requires an
explicit `read_only=False` rather than quietly writing to a database you
asked not to change.

Outlook and Publisher are not supported. Outlook is single-instance per
user and has no per-document VBA project, so owning it would mean owning
your mail client. Publisher exposes no VBA project on its documents and
retires in October 2026.

## Requirements

- Windows.
- The desktop application you want to drive. Tested against Microsoft 365
  x64; 2016 and newer expected to work.
- Python 3.10 or newer. Tested on 3.14.
- `pywin32`, installed as a dependency.

### Trust access to the VBA project object model

Module injection goes through the VBA project object model, which Excel,
Word and PowerPoint block by default. Enable it in each application you
intend to use, at:

```text
File > Options > Trust Center > Trust Center Settings > Macro Settings
  [x] Trust access to the VBA project object model
```

The first run fails with a message naming this setting if it is off. The
setting is per application and per Windows user: enabling it for Excel does
not affect Word, and enabling it for one account does not affect a service
account. Access has no such option, because its VBA project is always
reachable. While enabled, any code running as that user can modify VBA
projects, so it suits a development machine rather than a shared server.

Run `python -m pyvbaharness doctor` to verify this and the rest of the
environment. It reports every host it finds and the ones it does not, and
an application you have not installed is a warning rather than a failure.

## Install

```bash
pip install pyvbaharness
```

Optional extras: `pyvbaharness[fuzz]` adds Hypothesis for property-based
testing. For development, clone and `pip install -e ".[dev,fuzz]"`.

## Running VBA

`run_vba` injects source as a module and calls one procedure from it,
creating an unsaved in-memory workbook if none is open. `run_macro` calls a
procedure that already exists in the workbook.

```python
excel.run_vba(source, proc="Main", args=(1, 2), timeout=30)
excel.run_macro("Module1.Main", 1, 2, timeout=30)
```

Return values arrive as Python objects. Scalars, dates, and 1-D or 2-D
arrays convert directly; COM objects convert to a `"<object:Range>"` marker
so the run still completes.

`PyVbaLog` is injected into every workbook and collects output into
`result.output`. `Debug.Print` writes to an Immediate window nothing is
reading, and `MsgBox` opens a blocking dialog, so neither is usable under
automation.

```python
result = excel.run_vba("""
Public Sub Main()
    Dim i As Long
    For i = 1 To 3
        PyVbaLog "step " & i
    Next i
End Sub
""")
result.output  # ['step 1', 'step 2', 'step 3']
```

Single expressions can be evaluated without writing a procedure:

```python
excel.eval("WorksheetFunction.Sum(1, 2, 3)")  # 6.0
excel.eval('UCase$("abc")')                   # 'ABC'
```

## Outcomes

Every run reports one of five outcomes. `passed` and `vba-error` describe
the VBA under test; the other three describe the harness or the environment
and say nothing about whether the code is correct.

| Outcome | Meaning |
| --- | --- |
| `passed` | The procedure ran to completion |
| `vba-error` | VBA raised an error, captured in `result.error` |
| `timeout` | The run exceeded its deadline; Excel was terminated |
| `modal-blocked` | A dialog requiring a human decision appeared; Excel was terminated |
| `runner-error` | The harness or a COM call failed around the run |

## Errors

Errors are captured inside VBA, so no dialog opens. Injected source is
instrumented with `Erl` line numbers and a per-procedure handler, which
supplies the failing line and the call stack.

```python
result = excel.run_vba("""
Public Sub Main()
    Dim x As Long
    x = 1
    Err.Raise 513, "MyModule", "something went wrong"
End Sub
""")

result.outcome            # 'vba-error'
result.error.number       # 513
result.error.source       # 'MyModule'
result.error.description  # 'something went wrong'
result.error.line         # 5
result.error.stack        # [('PyVbaUserCode.Main', 5)]
```

For a nested failure, `error.stack` lists every frame the error passed
through, deepest first:

```python
[('Helpers.Parse', 12), ('Model.Load', 40), ('Main.Run', 7)]
```

`line_numbers=False` injects source unmodified, leaving `error.line` and
`error.stack` empty. Instrumentation is skipped for sources that already
contain numeric line labels, and an explicit `On Error` statement takes
precedence from the point it appears.

## Timeouts

Every command carries a deadline. On expiry the harness terminates its Excel
process, records the outcome as `timeout`, and starts a fresh instance for
the next call.

```python
result = excel.run_vba("""
Public Sub Main()
    Do
    Loop
End Sub
""", timeout=5.0)

result.outcome  # 'timeout'
```

For work of unpredictable duration, report progress from VBA and set
`idle_timeout`. Each report extends the deadline, so the run is terminated
only after it stops reporting.

```python
result = excel.run_vba(source, proc="Recalculate",
                       idle_timeout=60,
                       on_progress=lambda fraction, message: ...)
```

```vba
Public Sub Recalculate()
    For i = 1 To 10000
        ' ... work ...
        PyVbaProgress i / 10000, "row " & i
    Next i
End Sub
```

## Documents

Every host opens and creates documents through the same two methods.

```python
word.open_document(r"C:\reports\report.docm", read_only=True)
word.run_macro("Report.Rebuild", timeout=120)
word.save_as(r"C:\out\rebuilt.docm")
```

Documents open read-only by default and close without saving unless
`save_as` is called. `save_as` accepts only macro-enabled formats, because
the others drop the VBA project silently while alerts are suppressed.

Modules can be exported to and imported from `.bas` and `.cls` files on any
host:

```python
word.export_modules("vba/")   # VBIDE export, for version control
word.import_modules("vba/")   # document modules are skipped
```

## Workbooks and ranges (Excel)

```python
excel.open_workbook(r"C:\reports\model.xlsm", read_only=True)
excel.run_macro("Analysis.Recalculate", timeout=120)
rows = excel.read_range("Summary", "A1:D50")

excel.new_workbook()
excel.write_range("Sheet1", "A1", [[1, 2], [3, 4]])
excel.save_as(r"C:\out\result.xlsm")
```

`new_workbook` and `open_workbook` are Excel's names for `new_document` and
`open_document`; both spellings work. Range reads and writes transfer whole
blocks in one COM call. `reset_sheets()` clears all worksheets while keeping
injected modules, for use between tests. These four methods exist on
`ExcelSession` alone, since no other host has a grid.

## Testing

Zero-argument procedures named `Test*` are discovered and run individually.
`PyVbaAssert` and `PyVbaAssertEqual` produce structured failures; any other
error is reported with its number, line, and stack.

```python
results = excel.run_tests("""
Public Sub TestMath()
    PyVbaAssertEqual 4, 2 + 2
End Sub

Public Sub TestBroken()
    PyVbaAssertEqual 5, 2 + 2, "arithmetic is broken"
End Sub
""")

for case in results:
    print(case.name, case.passed, case.result.error)
```

A test that hangs is reported as `timeout`. The session recycles, the test
module is reinjected, and the remaining tests run. When recovery is not
possible, such as with `auto_recycle=False`, the remaining tests are
reported as not run.

### pytest integration

Procedures in files named `test_*.bas` are collected as pytest items:

```bash
pytest tests/vba/
pytest -k Discount -v
pytest --junitxml=results.xml
```

Failures report the assertion message, error line, VBA stack, and any
`PyVbaLog` output. One auto-recycling session serves the run; under
pytest-xdist each worker process gets its own.

### Coverage

```python
excel.add_module("Model", source, coverage=True)
excel.run_tests(test_source)

report = excel.coverage_report()
report.percent                      # 87.5
report.modules["model"].missed      # [42, 43, 51]
```

Coverage is opt-in per module because instrumented code runs slower. Hits
accumulate across runs until the module is replaced.

### Property-based testing

```python
from pyvbaharness.properties import check_vba_function

check_vba_function(excel, source, "Discount",
                   check=lambda args, value: 0 <= value <= args[0])
```

Input strategies are derived from the parsed VBA signature. Hypothesis
shrinks failures to a minimal counterexample. Requires
`pip install pyvbaharness[fuzz]`.

## Batch execution (Excel)

`run_batch` runs many calls in one COM round trip. It stages its arguments
on a hidden worksheet, so it is available on `ExcelSession` only. Results are returned in
call order with the same detail as `run_macro`, including per-call errors
with line and stack.

```python
results = excel.run_batch([
    ("Model.Score", (row, weight)) for row, weight in inputs
])
```

Measured against equivalent serial calls: 5.4x at 200 calls and 47x at
1000, where per-call cost falls to 0.062 ms. Arguments must be scalars.

## Parallel execution

`SessionPool` distributes work across several owned instances. Each member
is a full session with its own process, watchdogs, and recovery, so a hang
recycles one member while the others continue. Pass `app=` to pool a host
other than Excel; a PowerPoint pool larger than one member is refused,
because PowerPoint cannot produce a second process.

```python
from pyvbaharness import SessionPool

with SessionPool(4) as pool:
    futures = [pool.run_vba(source, proc="Crunch", args=(n,))
               for n in range(24)]
    results = [f.result() for f in futures]

    cases = pool.run_tests(suite_source, timeout=30)  # sharded across members

    future = pool.submit(lambda s: (                  # multi-step flow
        s.open_workbook(r"C:\data\model.xlsm"),
        s.run_macro("Model.Recalculate", timeout=300),
        s.read_range("Out", "A1:C10"),
    )[-1])
```

Throughput on a 16-core machine with 120 ms tasks: 2.0x at two members, 3.7x
at four, 4.8x at six (`benchmarks/output/pool-baseline-1.1.0.json`). Each
member uses 150 to 300 MB of RAM. Compile checks remain serialized
machine-wide inside a pool because they drive the visible VBE, which is a
shared surface; hidden runs and range IO do not interfere with each other.

## Compile checking

```python
result = excel.compile_project(watch_seconds=15)
if result.outcome == "rejected":
    print(result.dialog.message)
```

Excel is made visible for the duration of the check. A hidden Excel does not
surface the compile-error dialog, which would report a rejection as a pass.
A clean project usually returns in about a second: the VBE disables its
Compile command once compilation succeeds, and the harness watches for that
signal. An `infrastructure-failure` outcome means the check could not
complete and the result is unknown.

Source that calls `PyVbaLog` or the assertion helpers requires
`compile_project(include_harness_support=True)` so those names resolve as
they do at run time.

## Command line

```bash
pyvbaharness doctor --live
pyvbaharness run My.bas --proc Main --arg 42
pyvbaharness run model.xlsm --proc Model.Recalculate
pyvbaharness check My.bas Other.cls
pyvbaharness check model.xlsm
```

`python -m pyvbaharness <command>` works identically.

A positional path routes on its extension. An Excel extension is opened as a
workbook, so `check model.xlsm` and `check --workbook model.xlsm` do the same
thing, and `run model.xlsm` opens it read-only and calls `--proc` from it.
Anything else is read as VBA source and injected into a fresh workbook. A
Word, PowerPoint, or Access document is refused with a message naming the
session class to use instead, since the command line drives Excel only.

Source files are read as UTF-8, falling back to the ANSI code page, because
that is what VBIDE export writes.

`doctor` checks each installed Office application, pywin32, the VBA project
trust setting, and the VBE error-trapping mode; "Break on All Errors" sends
handled errors to the debugger and stalls automation. `--live` additionally
starts an owned Excel and runs a smoke test.

Exit codes: `0` pass or accepted, `1` VBA failure or rejected compile, `2`
infrastructure failure.

## Dialog handling

A watcher thread scans the owned Excel process for dialogs while a run is in
flight and applies a deliberately narrow policy:

| Dialog | Action |
| --- | --- |
| VBA runtime error | Dismissed with End |
| Informational (OK, or OK and Help only) | Dismissed with OK |
| Anything with a real choice (Yes/No, Cancel, Retry, Debug, Save) | Reported as `modal-blocked`; Excel terminated |
| Compile error, outside a compile check | Reported as `modal-blocked`; Excel terminated |
| Excel's own prompts, such as "save your changes?" | Reported as `modal-blocked`; Excel terminated |

Excel's own prompts are not classic Win32 dialogs. "Do you want to save your
changes?" is a `NUIDialog` whose controls are drawn inside a NetUI surface,
with no Win32 buttons to enumerate or click, so its text cannot be read and
it is never dismissed. It is detected by modality instead: a titled Excel
window is disabled for as long as something modal owns the application,
whatever class that something is. The result is `modal-blocked` rather than
a bare timeout.

Prompts of this kind are suppressed in the first place by `DisplayAlerts`,
`AskToUpdateLinks`, `UpdateLinks:=0`, `IgnoreReadOnlyRecommended`, and
closing without saving. Detection covers the case where VBA under test turns
suppression back on. Alert suppression is re-asserted before the harness
closes, saves, or quits, so user code cannot leave a prompt waiting for the
teardown path.

`result.dialogs` records what appeared and what was done about it. On a
blocked dialog or a timeout, a screenshot of the Excel window is captured
where possible and referenced from the result, which is the practical way to
see an Excel prompt whose text cannot be read.

## Avoiding wedges

Prompts that can be prevented are prevented rather than dismissed, because
Office's own dialogs draw their controls inside a NetUI surface with no
Win32 buttons: the harness can see them and cannot click them. Alerts, link
prompts, AutoRecover and feature-install prompts are all turned off before a
document exists, and each host's teardown removes what would otherwise raise
a save prompt.

Where a prompt cannot be prevented, the harness prefers a signal the system
reports exactly over a sampled guess: it waits on the process handle for
exit rather than polling, treats the VBE disabling its Compile control as
proof a compile finished, and holds its modal check exactly as long as a
clicked dialog still exists. Deadlines remain as the backstop for the one
case no signal covers, a COM call into a blocked apartment, which cannot be
interrupted from inside the process making it.

## Process ownership

The harness creates its own Excel instance and never attaches to a running
one. It verifies this by comparing `EXCEL.EXE` process IDs before and after
creation, and refuses to proceed if the instance already existed, since a
timeout would terminate it.

The owned instance is placed in a kernel job object with kill-on-close, so
Windows terminates Excel if the harness process dies for any reason,
including a force kill. Manifest files provide a second layer: they record
process ID and start time together, so a later sweep cannot terminate an
unrelated process that reused the ID.

One session runs at a time per machine by default, preventing accidental
overlap. `SessionPool` is the supported route to concurrency.

Run targets must be a plain `Proc` or `Module.Proc` inside the harness
workbook. Workbook-qualified targets are rejected before any COM call.

## Performance

Measured on Excel 365 x64 with Python 3.14
(`benchmarks/output/baseline-1.1.0.json`):

| Operation | Median |
| --- | --- |
| Session startup and teardown | 3.0 s (0.6 s to start, 2.4 s for Excel to exit) |
| Run a procedure, same target as previous call | 0.5 ms |
| Run a procedure with arguments | 0.6 ms |
| `run_vba` with unchanged source | 0.7 ms |
| Run a different target (dispatcher regenerated) | 77 ms |
| Batched calls, 1000 per batch | 0.062 ms each |
| Compile check, clean project | 0.9 s |
| Write 10,000 cells | 66 ms |
| Read 10,000 cells | 8 ms |

Most of a session's lifetime cost is Excel exiting, not starting: Quit
returns well before the process does, and teardown waits on the process
handle to confirm termination rather than assuming it. Reuse a session
across runs, or a `SessionPool`, if that matters.

Per-run cost depends on three caches: resolved target signatures, injected
source, and the generated dispatcher. Repeated calls to the same target with
unchanged source hit all three. Changing the target regenerates the
dispatcher, which accounts for the 77 ms figure.

## Development

```bash
python -m pytest tests/unit                          # 208 tests, no Office
python -m pytest tests/live -m live -o addopts=""    # 120 tests, real Office
python benchmarks/run_benchmarks.py
python benchmarks/run_pool_benchmarks.py
```

The unit suite covers dialog policy, trace validation, signature parsing,
code generation, source instrumentation, write chunking and the per-app
capability table, plus the supervisor state machine driven by a fake worker
that speaks the same pipe protocol. The live suite covers real behavior on
all four hosts, including deliberate hangs, blocking dialogs, and a worker
terminated mid-run; it asserts that no Office processes survive. Running it
opens and closes real applications, and the PowerPoint tests put a window on
screen.

## Documentation

| Document | Contents |
| --- | --- |
| [Architecture](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/docs/architecture.md) | Process model, hang-resistance layers, design rationale |
| [Troubleshooting](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/docs/troubleshooting.md) | What each failure means and what to do about it |
| [Implementation guide](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/docs/IMPLEMENTATION_GUIDE.md) | How to change the code; catalog of measured Office behaviors |
| [Releasing](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/docs/RELEASING.md) | Version bump, validation, and the PyPI publishing workflow |

## License

MIT. See [LICENSE](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/LICENSE).
