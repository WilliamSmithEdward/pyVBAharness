# pyVBAharness

Run VBA inside real desktop Excel from Python and always get your process
back.

Driving Excel from Python has one characteristic failure: your script stops
and never resumes. A VBA runtime error opens a modal dialog that waits for
someone to click it. `Application.Run` accepts no timeout. Excel is not a
child of your process either, so killing the script leaves a hidden
EXCEL.EXE behind, still holding the workbook.

This harness deals with all three. Every run sits behind a watchdog in a
separate supervisor process that holds no COM references of its own. VBA
errors are trapped inside VBA and come back as data. When something wedges,
the exact Excel process the harness started is killed, and your next call
gets a fresh one.

```python
from pyvbaharness import ExcelSession

with ExcelSession() as excel:
    result = excel.run_vba("""
Public Sub Main()
    PyVbaLog "hello from VBA"
End Sub
""")
    print(result.outcome)  # passed
    print(result.output)   # ['hello from VBA']
```

## Requirements

- Windows with desktop Microsoft Excel. Tested against Microsoft 365 x64;
  2016 and newer should work.
- Python 3.10 or newer. Tested on 3.14.
- `pywin32`, installed as a dependency.

### The Excel setting you have to enable

The harness injects VBA modules through the VBA project object model, and
Excel blocks that access until you allow it. Open File > Options > Trust
Center > Trust Center Settings > Macro Settings, and tick **Trust access to
the VBA project object model**.

Skip this and the first run fails with a message naming the setting. It
applies per Office application and per Windows user, so enabling it in Excel
leaves Word alone, and enabling it for you does nothing for a service
account. Any code on the machine can modify VBA projects once it is on, which
makes a development machine the right place for it and a shared server the
wrong one.

`python -m pyvbaharness doctor` checks this and the rest of the environment.

## Install

```bash
pip install -e .
```

## Using it

### Run some VBA

`run_vba` injects the source as a module and calls one procedure from it. If
no workbook is open, it makes an unsaved one in memory first.

```python
from pyvbaharness import ExcelSession

with ExcelSession() as excel:
    result = excel.run_vba("""
Public Function AddNums(ByVal a As Long, ByVal b As Long) As Long
    AddNums = a + b
End Function
""", proc="AddNums", args=(20, 22))

    assert result.outcome == "passed"
    assert result.value == 42
```

Scalars, dates, and 1-D or 2-D arrays return as Python values. An object
comes back as a `"<object:Range>"` marker so the run still completes; return
a scalar or write to cells if you need the data.

### Collect output

The harness injects a `PyVbaLog` sub. Whatever you pass it lands in
`result.output`. `Debug.Print` needs an Immediate window nobody is watching,
and `MsgBox` opens a dialog that blocks the run, so this is the one that
works under automation.

```python
result = excel.run_vba("""
Public Sub Main()
    Dim i As Long
    For i = 1 To 3
        PyVbaLog "step " & i
    Next i
End Sub
""")
print(result.output)  # ['step 1', 'step 2', 'step 3']
```

### Read VBA errors, with line numbers and a stack

Errors are caught inside VBA and returned as data, so no dialog ever opens.
On the way in, the source gets classic `Erl` line numbering plus a small
handler in each procedure, which is what makes the line and the stack
available.

```python
result = excel.run_vba("""
Public Sub Main()
    Dim x As Long
    x = 1
    Err.Raise 513, "MyModule", "something went wrong"
End Sub
""")

result.outcome           # 'vba-error'
result.error.number      # 513
result.error.source      # 'MyModule'
result.error.description # 'something went wrong'
result.error.line        # 5, the Err.Raise line above
result.error.stack       # [('PyVbaUserCode.Main', 5)]
```

When the failure is nested, `error.stack` gives you the whole chain,
deepest frame first, each with its own line:

```python
[('Helpers.Parse', 12), ('Model.Load', 40), ('Main.Run', 7)]
```

Pass `line_numbers=False` to inject the source untouched; `error.line` and
`error.stack` go empty. Sources that already carry numeric labels are left
alone, and your own `On Error` statement takes over wherever you write one.

### Batch many calls into one round trip

```python
results = excel.run_batch([
    ("Model.Score", (row, weight)) for row, weight in inputs
])
```

Each result carries everything `run_macro` gives you: value, output, and an
error with its line and stack. The difference is that the COM round trip is
paid once for the whole batch instead of once per call. Measured against
serial calls: 2.7x at 50 calls, 4.7x at 200, 9.9x at 3000, where per-call
cost bottoms out around 0.066 ms. Arguments have to be scalars.

### Let a long run prove it is still alive

A 20 minute recalculation should not need you to guess a 25 minute timeout.
Have the VBA report progress, then set an idle timeout. Each report pushes
the deadline forward, so the run is only killed once it goes quiet.

```python
result = excel.run_vba(source, proc="Recalculate",
                       idle_timeout=60,
                       on_progress=lambda pct, msg: print(pct, msg))
```

```vba
Public Sub Recalculate()
    For i = 1 To 10000
        ' ... work ...
        PyVbaProgress i / 10000, "row " & i
    Next i
End Sub
```

### Measure line coverage

```python
excel.add_module("Model", source, coverage=True)
excel.run_tests(test_source)
report = excel.coverage_report()
print(report.percent)                      # 87.5
print(report.modules["model"].missed)      # [42, 43, 51]
```

Coverage is per module and opt-in, since instrumented code runs slower. Hits
accumulate across runs until the module changes.

### Evaluate an expression

```python
excel.eval("WorksheetFunction.Sum(1, 2, 3)")  # 6.0
excel.eval('UCase$("abc")')                   # 'ABC'
```

### Run VBA tests

Zero-argument procedures named `Test*` are discovered and run one at a time.
`PyVbaAssert` and `PyVbaAssertEqual` produce structured failures; anything
else that goes wrong is reported with its error number and line.

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
    print(case.name, "passed" if case.passed else case.result.error)
```

One hung test costs you that test. The session recycles, the module is
reinjected, and the rest of the suite still runs. Where recovery is
impossible, say with `auto_recycle` switched off, the remaining tests are
reported as not run so nothing disappears quietly.

### Or run them through pytest

Name a file `test_*.bas` and each zero-argument `Test*` procedure inside
becomes a pytest item:

```bash
pytest tests/vba/            # collects test_model.bas
pytest -k Discount -v        # selection works
pytest --junitxml=out.xml    # so does CI reporting
```

Failures show the assertion message, the error line, the VBA stack, and any
`PyVbaLog` output. A single auto-recycling session serves the whole run.
Under pytest-xdist each worker gets its own.

### Fuzz a function

```python
from pyvbaharness.properties import check_vba_function

check_vba_function(excel, source, "Discount",
                   check=lambda args, value: 0 <= value <= args[0])
```

Input strategies come from the parsed VBA signature, and Hypothesis shrinks
any failure down to a minimal counterexample. Install with
`pip install pyvbaharness[fuzz]`.

### Survive a hang

```python
result = excel.run_vba("""
Public Sub Main()
    Do
    Loop
End Sub
""", timeout=5.0)

result.outcome  # 'timeout'
```

Excel was killed and the session recycled itself, so the next call runs
normally against a fresh instance. A timeout says only that the run did not
finish. It is never reported as a pass or a failure of your code.

### Work with workbooks and cells

```python
excel.open_workbook(r"C:\reports\model.xlsm", read_only=True)
excel.run_macro("Analysis.Recalculate", timeout=120)
rows = excel.read_range("Summary", "A1:D50")

excel.new_workbook()
excel.write_range("Sheet1", "A1", [[1, 2], [3, 4]])
excel.save_as(r"C:\out\result.xlsm")
```

Workbooks open read-only by default and close without saving unless you call
`save_as`. Ranges move as whole blocks in a single call, never cell by cell.

### Check that a project compiles

```python
result = excel.compile_project(watch_seconds=15)
if result.outcome == "rejected":
    print(result.dialog.message)  # the VBE's compile error text
```

Excel goes visible for the duration. A hidden Excel never surfaces the
compile-error dialog, which would quietly turn a rejection into a false pass.
A clean project usually answers in about a second, because the VBE disables
its Compile command once compilation succeeds and the harness watches for
that. The watch window only runs out when no signal arrives at all, and then
you get `infrastructure-failure`, meaning the verdict is unknown.

Code that calls `PyVbaLog` or the assert helpers needs
`compile_project(include_harness_support=True)`, so those names resolve the
same way they will at run time.

### Run in parallel

`SessionPool` spreads work across several owned Excel instances. Every member
is a complete session with its own process, watchdogs, and recovery, so a
hang costs that one member a recycle while the rest keep working.

```python
from pyvbaharness import SessionPool

with SessionPool(4) as pool:
    futures = [pool.run_vba(source, proc="Crunch", args=(n,))
               for n in range(24)]
    results = [f.result() for f in futures]

    # Shard a VBA test suite across all members:
    cases = pool.run_tests(suite_source, timeout=30)

    # Or borrow a member for a multi-step flow:
    future = pool.submit(lambda s: (
        s.open_workbook(r"C:\data\model.xlsm"),
        s.run_macro("Model.Recalculate", timeout=300),
        s.read_range("Out", "A1:C10"),
    )[-1])
```

On a 16-core machine running 120 ms tasks, throughput scales 1.9x at two
members, 3.4x at four, and 4.3x at six
(`benchmarks/output/pool-baseline-0.4.0.json`). Budget 150 to 300 MB of RAM
per member. Compile checks stay serialized across the whole machine even
inside a pool, because they drive the visible VBE and that surface really is
shared. Hidden runs and range IO overlap without interfering.

### Command line

```bash
python -m pyvbaharness doctor --live   # diagnose the environment
python -m pyvbaharness run My.bas --proc Main --arg 42
python -m pyvbaharness check My.bas Other.cls
python -m pyvbaharness check --workbook model.xlsm
```

`doctor` looks at Excel, pywin32, the VBA project trust setting, and the VBE
error-trapping mode, where "Break on All Errors" would drop every handled
error into the debugger and stall automation. Add `--live` and it starts an
owned Excel for a smoke run. Exit codes are consistent across commands: 0 for
pass or accepted, 1 for a VBA failure or a rejected compile, 2 for anything
infrastructural.

## Outcomes

Every run reports one of five outcomes. The first two describe your VBA. The
rest describe the harness or the environment.

| Outcome | Meaning |
| --- | --- |
| `passed` | The procedure ran to completion. |
| `vba-error` | VBA raised an error, captured in `result.error`. |
| `timeout` | The run passed its deadline and Excel was killed. |
| `modal-blocked` | A dialog needing a human decision appeared; Excel was killed. |
| `runner-error` | The harness or COM failed around the run. |

## Dialogs

A watcher thread scans the owned Excel process for dialogs the whole time a
run is in flight, and applies a deliberately narrow policy:

- A VBA runtime-error dialog is dismissed with End.
- A dialog with nothing but OK, or OK and Help, is dismissed with OK, so a
  stray `MsgBox` cannot stall the run.
- Anything offering a real choice, such as Yes/No, Cancel, Retry, Debug, or
  Save changes, is never guessed at. The run reports `modal-blocked` and
  Excel is killed.
- Compile-error dialogs are left alone outside a compile check.

`result.dialogs` records what appeared and what was done about it.

## Safety

The harness starts its own Excel and never attaches to one you already have
open. It confirms this at startup by comparing process IDs before and after,
and refuses to run if the instance turns out to be pre-existing, since a
timeout would kill it.

Only one session runs at a time per machine by default. That lock is there to
prevent accidental overlap; when you want concurrency, `SessionPool` provides
it with proper isolation and keeps compile checks serialized.

The owned Excel goes into a kernel job object with kill-on-close, so if the
harness process dies for any reason, including a force kill, Windows
terminates Excel along with it. Manifest sweeps back that up, matching on
process ID and start time together so a recycled ID can never send the sweep
after an unrelated process.

Run targets have to be a plain `Proc` or `Module.Proc` inside the harness
workbook. Workbook-qualified targets are rejected before any COM call
happens.

## Tests

```bash
python -m pytest tests/unit
```

147 tests, no Excel required. Dialog policy, trace validation, VBA signature
parsing, code generation, source instrumentation, write chunking, and the
supervisor state machine all run against a scripted fake worker that speaks
the same pipe protocol.

```bash
python -m pytest tests/live -m live -o addopts=""
```

54 tests against real Excel, including deliberate hangs, blocking dialogs,
and a worker killed mid-run. The suite leaves no Excel processes behind, and
it fails if it does.

```bash
python benchmarks/run_benchmarks.py
python benchmarks/run_pool_benchmarks.py
```

## Performance

Measured on Excel 365 x64 with Python 3.14
(`benchmarks/output/baseline-0.4.0.json`):

| Operation | Median |
| --- | --- |
| Session startup and teardown | 0.5 s warm, around 3 s cold |
| Run a procedure, same target as last time | 0.6 ms |
| Run a procedure with arguments | 1.4 ms |
| Repeat `run_vba` with unchanged source | 1.9 ms |
| Run a different target, regenerating the dispatcher | 97 ms |
| Batched calls, 1000 at a time | 0.107 ms each |
| Compile check on a clean project | 1.2 s |
| Write 10,000 cells | 71 ms |
| Read 10,000 cells | 10 ms |

Holding one session open is what buys the speed. A fresh session per run
costs seconds; a warm run costs well under a millisecond. Three caches keep
it there: resolved target signatures, injected source, and the generated
dispatcher. Loop over `run_vba` with the same source and you stay on the fast
path. Once the calls themselves are trivially short, `run_batch` takes out
what remains.

## Documentation

- [Architecture](docs/architecture.md) covers the process model, the layers
  of hang resistance, and why each decision was made.
- [Troubleshooting](docs/troubleshooting.md) explains what each failure means
  and what to do next.
- [Implementation guide](docs/IMPLEMENTATION_GUIDE.md) is for changing the
  code, and includes the catalog of measured Excel behaviors this harness
  works around.

## License

MIT. See [LICENSE](LICENSE).
