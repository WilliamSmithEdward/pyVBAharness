# pyVBAharness

Run VBA inside real desktop Excel from Python, without the run ever hanging
your process.

Excel automation fails in ways ordinary code does not: a VBA runtime error
opens a modal dialog and waits forever, `Application.Run` blocks with no
timeout, and the Excel process it started is not a child process, so killing
your script leaves Excel running. pyVBAharness puts every run behind a
watchdog, captures VBA errors inside VBA so they never reach a dialog, and
kills the exact Excel process it owns when something wedges.

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

- Windows with desktop Microsoft Excel installed (Microsoft 365 or 2016+)
- Python 3.10 or newer
- `pywin32` (installed automatically)

### Required Excel setting

**pyVBAharness needs "Trust access to the VBA project object model" turned
on.** It injects VBA modules through the VBA project object model, and Excel
blocks that access by default.

To enable it: **File > Options > Trust Center > Trust Center Settings >
Macro Settings**, then tick **Trust access to the VBA project object model**.

Without it, the first run fails with a message naming this setting. The
option applies per Office application and per user. It lets any code running
on your machine modify VBA projects, so enable it on a development machine
rather than a shared or production one.

## Install

```bash
pip install -e .
```

## Using it

### Run some VBA

`run_vba` injects the source as a module and runs one procedure from it. If
no workbook is open, it creates an unsaved one in memory.

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

`result.value` carries the procedure's return value. Scalars, dates, and 1-D
or 2-D arrays come back as Python values; objects come back as a
`"<object:Range>"` marker rather than failing the run.

### Collect output from VBA

The harness injects a `PyVbaLog` sub. Anything your code passes to it is
returned in `result.output`, which is more reliable than `Debug.Print` (no
Immediate window to read) and safer than `MsgBox` (no modal dialog).

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

### Handle VBA errors, with line numbers

A VBA error is captured inside VBA and returned as data. It never opens a
dialog and never stops the run. The source is instrumented on the way in
(classic ``Erl`` line numbering plus a per-procedure handler), so the error
also reports which line raised it.

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
result.error.line        # 5 (the Err.Raise line in the source above)
```

Pass ``line_numbers=False`` to inject the source verbatim; ``error.line`` is
then None. The instrumentation skips sources that already carry numeric
labels, and a user ``On Error`` statement overrides it from that point on.

### Evaluate an expression

```python
excel.eval("WorksheetFunction.Sum(1, 2, 3)")  # 6.0
excel.eval('UCase$("abc")')                   # 'ABC'
```

### Run VBA tests

Procedures named ``Test*`` (zero arguments) are discovered and run
individually. ``PyVbaAssert`` and ``PyVbaAssertEqual`` raise structured
failures; anything else that errors is reported with its number and line.

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

If a test hangs, it reports `timeout`, the session recycles, the test module
is reinjected, and the remaining tests still run; a hang costs one test, not
the suite. (Only when recovery is impossible, for example `auto_recycle`
disabled, is the remainder reported as not run rather than silently
dropped.)

### Survive a hang

```python
result = excel.run_vba("""
Public Sub Main()
    Do
    Loop
End Sub
""", timeout=5.0)

result.outcome  # 'timeout': Excel was killed, the session recycled itself
```

The next call works normally: the session starts a fresh Excel. Note that a
timeout tells you nothing about your VBA except that it did not finish; it is
never treated as a pass or a failure of the code.

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
`save_as`. Reads and writes move whole blocks in one call rather than
cell by cell.

### Check that a project compiles

```python
result = excel.compile_project(watch_seconds=15)
if result.outcome == "rejected":
    print(result.dialog.message)  # the VBE's compile error text
```

Excel is made visible for the duration of a compile check: a hidden Excel
does not surface the compile-error dialog, which would silently turn a
rejection into a false pass. A clean compile returns in about a second (the
VBE disables its Compile command once the project compiles, which is a
positive done-signal); the watch window is only waited out when no signal
arrives. An `infrastructure-failure` outcome means the check could not be
completed and the verdict is unknown.

Code that calls `PyVbaLog` or the assert helpers needs
`compile_project(include_harness_support=True)` so those names resolve the
way they do during a run.

### Run in parallel

`SessionPool` runs work across several owned Excel instances at once. Each
member is a full session (own Excel process, own watchdogs, own recovery);
a hang costs that member one recycle while the others keep working.

```python
from pyvbaharness import SessionPool

with SessionPool(4) as pool:
    futures = [pool.run_vba(source, proc="Crunch", args=(n,))
               for n in range(24)]
    results = [f.result() for f in futures]

    # Shard a VBA test suite across all members:
    cases = pool.run_tests(suite_source, timeout=30)

    # Or borrow a session for a multi-step flow:
    future = pool.submit(lambda s: (
        s.open_workbook(r"C:\data\model.xlsm"),
        s.run_macro("Model.Recalculate", timeout=300),
        s.read_range("Out", "A1:C10"),
    )[-1])
```

Measured scaling on a 16-core machine with ~120 ms tasks
(`benchmarks/output/pool-baseline-0.3.0.json`): 1.7x at 2 members, 2.3x at
4, flat after that; longer-running tasks scale closer to linearly because
the fixed ~20 ms per-run overhead amortizes. Budget 150-300 MB RAM per
member. Compile checks stay serialized machine-wide even inside a pool (they
drive the visible VBE, which is a genuinely shared surface); hidden runs and
range IO parallelize cleanly.

### Command line

```bash
python -m pyvbaharness doctor --live   # diagnose the environment
python -m pyvbaharness run My.bas --proc Main --arg 42
python -m pyvbaharness check My.bas Other.cls
python -m pyvbaharness check --workbook model.xlsm
```

`doctor` verifies Excel, pywin32, the VBA-project trust setting, and the VBE
error-trapping mode ("Break on All Errors" would stop every handled error in
the debugger), and with `--live` starts an owned Excel for a smoke run. Exit
codes everywhere: 0 pass/accepted, 1 VBA failure/rejected, 2 infrastructure.

## Outcomes

Every run reports one of five literal outcomes. Only the first two say
anything about your VBA:

| Outcome | Meaning |
| --- | --- |
| `passed` | The procedure ran to completion. |
| `vba-error` | VBA raised an error, captured in `result.error`. |
| `timeout` | The run exceeded its deadline; Excel was killed. |
| `modal-blocked` | A dialog needing a human decision appeared; Excel was killed. |
| `runner-error` | The harness or COM failed before or around the run. |

## Dialogs

Dialogs are watched for continuously and handled by a conservative policy:

- A VBA runtime-error dialog is dismissed with **End**.
- A purely informational dialog (only OK, or OK and Help) is dismissed
  with **OK**, so a stray `MsgBox` does not stall the run.
- Anything offering a real choice (Yes/No, Cancel, Retry, Debug, Save?) is
  never guessed at. The run reports `modal-blocked` and Excel is killed.
- A compile-error dialog is never dismissed outside a compile check.

`result.dialogs` lists what was seen and what was done about it.

## Safety

- The harness always creates its own Excel process and never attaches to one
  you already have open. It verifies this at startup and refuses to run if
  the instance turns out to be pre-existing, because a timeout kills that
  process.
- Only one harness session runs at a time per machine by default; the lock
  exists to stop accidental concurrency. Deliberate concurrency goes through
  `SessionPool`, which isolates members per Excel process and keeps the one
  genuinely shared operation (compile checks) serialized.
- The owned Excel is placed in a kernel job object with kill-on-close: if
  the harness process dies for any reason, including being force-killed, the
  kernel terminates Excel with it. Manifest-based sweeps remain as backup,
  matching on both process id and start time so they cannot kill an
  unrelated process that reused the id.
- Run targets must be a plain `Proc` or `Module.Proc` inside the harness
  workbook. Workbook-qualified targets are rejected before any COM call.

## Tests

```bash
python -m pytest tests/unit
```

The unit suite needs no Excel: dialog policy, trace validation, VBA signature
parsing, code generation, write chunking, and the supervisor state machine
(driven by a scripted fake worker) are all covered without COM.

```bash
python -m pytest tests/live -m live -o addopts=""
```

The live suite drives real Excel, including deliberate hangs and blocking
dialogs. It leaves no Excel processes behind.

```bash
python benchmarks/run_benchmarks.py
```

## Performance

Measured on Excel 365 x64, Python 3.14 (`benchmarks/output/baseline-0.2.0.json`):

| Operation | Median |
| --- | --- |
| Session startup and teardown (new Excel process) | 0.5 s warm, ~3 s cold |
| Run a procedure, same target as last time | 15 ms |
| Repeat `run_vba` with identical source (injection cache) | 20 ms |
| Run a procedure, different target | 76 ms |
| Compile check, clean project | 1.0 s |
| Write 10,000 cells | 67 ms |
| Read 10,000 cells | 9 ms |

Keeping one session open across many runs is what makes this fast: each
`run_vba(...)` in a loop costs milliseconds, while a fresh session per run
costs seconds. Identical source is never reinjected, so calling `run_vba`
in a loop is as cheap as `run_macro`.

## Documentation

- [Architecture](docs/architecture.md): process model, hang-resistance
  layers, and the design decisions behind them
- [Troubleshooting](docs/troubleshooting.md): what each failure means and
  what to do about it

## License

MIT. See [LICENSE](LICENSE).
