# pyVBAharness Implementation Guide

How to work on this codebase without breaking the guarantees it exists to
provide. Written for coding agents and engineers making changes; read it
before editing anything under `src/pyvbaharness/`.

Companion documents: [architecture.md](architecture.md) explains what the
design is and why; [troubleshooting.md](troubleshooting.md) explains what
failures mean for users; [AGENTS.md](../AGENTS.md) states the operating
rules. This guide is the how-to.

## 1. Orientation in five minutes

The harness runs VBA inside real desktop Excel and refuses to hang. It is a
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
| `worker/excel_host.py` | every COM call in the project |
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

## 3. The Excel behavior catalog

Every entry below was measured on Excel 365 x64, cost real debugging time,
and has a live test guarding it. Treat this as ground truth. Do not "clean
up" code that references these; the workaround is the feature.

### 3.1 Dispatch attaches instead of creating

`win32com.client.Dispatch("Excel.Application")` calls `GetActiveObject`
first and returns whatever Excel is already running. Since the harness
kills its Excel on a hang, attaching puts a user's open workbooks in the
blast radius. Use `CoCreateInstance`, and keep the PID-snapshot check in
`ExcelHost.create` that refuses to proceed if the "new" instance already
existed.

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

### 3.11 Excel is not a child of its COM client

DCOM launches it, so killing the worker tree never reaps Excel. The harness
records the PID (from `GetWindowThreadProcessId(app.Hwnd)`) plus its start
time, and additionally places Excel in a kill-on-close job object so worker
death of any kind takes Excel with it.

## 4. Invariants

These are the guarantees the harness sells. Changing one is a contract
change: update the docs, the oracle, and the tests in the same patch.

1. The harness creates its own Excel and never attaches to a running one.
2. Every command carries a positive timeout. A breach kills the recorded
   Excel, kills the worker, and marks the session dead.
3. No command runs after the owned Excel is killed.
4. A timeout or blocked modal is infrastructure state, never evidence about
   the VBA under test. Only `passed` and `vba-error` describe the code.
5. VBA errors are captured inside VBA by a directly-called dispatcher.
6. Dialogs offering a real choice are never dismissed by guessing.
7. All COM lives in the worker; supervisor and watcher stay COM-free.
8. Compile checks serialize machine-wide (they drive the visible VBE).
9. One session serves one caller at a time; `SessionPool` enforces this by
   checkout.
10. Workbooks open read-only by default and close without saving.

`oracle.py` encodes most of these as a trace validator. Sessions validate
their own trace on close and warn, and the unit suite replays synthetic
traces, so a regression in safety behavior fails loudly rather than
silently.

## 5. Recipe: add a new capability end to end

Worked example, mirroring how `reset_sheets` was added.

**Step 1, name the command.** Add `CMD_RESET_SHEETS = "reset_sheets"` to
`protocol.py`. Add an event kind too if the feature reports asynchronously.

**Step 2, implement the COM work** in `worker/excel_host.py`, decorated so
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
`benchmarks/output/baseline-0.4.0.json`):

| Operation | Cost |
| --- | --- |
| Session startup and teardown | 0.5 s warm |
| Warm run, same target | 0.6 ms |
| Run with arguments | 1.4 ms |
| `run_vba` with identical source (cache hit) | 1.9 ms |
| Retarget (dispatcher regenerated) | 97 ms |
| Batch, 1000 calls | 0.107 ms per call (7.5x) |
| Compile check, clean project | 1.2 s |
| Write 10,000 cells | 71 ms |
| Read 10,000 cells | 10 ms |

Where the speed comes from, so you do not accidentally remove it:

- **Signature cache** (`ExcelHost._resolved`): resolving a target used to
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
