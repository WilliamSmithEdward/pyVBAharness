# pyVBAharness Architecture

pyVBAharness is a Python harness that runs VBA inside real desktop Excel with
three explicit goals, in priority order when they conflict: hang resistance,
accuracy, performance. This document records the design, the reasoning, and the
provenance of each load-bearing decision.

## Provenance

The design distills two working implementations:

- ROneCOne (`tools/run_excel_tests.ps1`, `run_excel_tests_worker.ps1`,
  `watch_vbe_dialogs.ps1`): supervisor/worker split, owned-process manifest,
  conservative dialog dismissal, break-mode termination, staged error
  reporting, exact-value validation anchors.
- XLIDE (`assets/testhost/run-vba-tests.ps1`, `XlideTestModalWatcher.cs`,
  `src/vbaTestHostSession.ts`, `src/vbaTestHostOracle.ts`,
  `src/vbaTestRunnerModuleCodegen.ts`, `syntax_corpus/oracle/*`): structured
  event protocol, three-watchdog session state machine, single abort path,
  in-VBA error capture returning JSON through `Application.Run`, modal
  classification policy, trace oracle, and the syntax-oracle reliability
  lessons (visible-Excel requirement for VBE dialogs, cold-compile latency,
  timeout-is-not-evidence, sequential-only automation).

## Process model

```text
caller (your code / pytest)
  -> ExcelSession (supervisor, pure Python, no COM)
       - spawns worker: python -m pyvbaharness.worker
       - sends JSON commands on worker stdin
       - consumes JSON events from worker stdout (prefix PYVBA_EVENT|)
       - enforces startup / per-command / cleanup watchdogs
       - kills by recorded PID on breach (taskkill /PID <pid> /T /F)
  -> worker process (all COM, STA main thread)
       - owns exactly one new Excel instance (DispatchEx, never attach)
       - records Excel PID from GetWindowThreadProcessId(app.Hwnd)
       - writes an owned-process manifest (pids + creation times)
       - runs a dialog-watcher daemon thread (ctypes user32 only)
  -> Excel process (EXCEL.EXE, hidden by default)
       - injected runner module traps VBA errors and returns JSON
```

Why a separate worker process: COM calls into Excel (`Workbooks.Open`,
`Application.Run`, `Close`, `Quit`, even COM release) are synchronous and can
block forever. A deadline can only be enforced from outside the COM apartment.
Both references landed on the same shape independently.

Why kill by recorded PID: Excel launched via COM is not a child of the worker
(DCOM launches it), so killing the worker tree does not kill Excel. The worker
therefore discovers the Excel PID from the Application window handle and
records it, with its process creation time, in a manifest the supervisor can
act on even if the worker itself is hung. PID + creation time guards against
PID reuse. Stale manifests from crashed runs are swept at session start; only
exact pid+ctime matches are killed.

## Hang-resistance layers

1. In-VBA error trapping. Every managed run goes through a generated
   dispatcher (`PyVbaHarnessCall`) that calls the target **directly**, wraps
   it in `On Error`, captures `Err.Number / Source / Description` plus
   collected log output, and returns one JSON string as the
   `Application.Run` return value. A VBA runtime error therefore never
   raises a dialog and never depends on window scraping to be reported.
   The call must be direct: see "Direct calls" below.
2. Dialog watcher. A daemon thread in the worker scans visible top-level
   windows of the owned Excel PID (250 ms period). `#32770` dialogs are
   captured (title, texts, buttons with control IDs) and classified:
   - `compile-error`: never dismissed; reported blocked (kill path).
   - `runtime-error` / `vba-modal`: dismissed with the `End` button, or `OK`
     when no decision buttons exist.
   - informational (only OK/Help buttons): dismissed with `OK`.
   - anything with decision buttons (Yes/No/Cancel/Retry/Debug/...): never
     guessed at; reported blocked (kill path).
   Clicks use `SendMessageTimeout` with `SMTO_ABORTIFHUNG` (500 ms) so a hung
   window cannot hang the watcher. The VBE main window becoming visible during
   a hidden run is reported and treated as blocked (debugger break).
3. Supervisor watchdogs. Startup watchdog (worker spawn to ready), per-command
   watchdog (every command carries a positive timeout), cleanup-grace watchdog
   (default 5 s) covering `Close`/`Quit`/COM release, which can themselves
   hang after the useful work finished. Breach follows one abort path:
   synthesize the terminal result, `taskkill /PID <excel> /T /F`, kill the
   worker, mark the session dead. A dead session never issues further
   commands (oracle rule: no macros after kill).
4. Kernel job object. The worker places its owned Excel in a job with
   kill-on-close: if the worker dies for ANY reason (crash, taskkill, the
   whole Python tree terminated), the kernel closes the job handle and kills
   Excel with no cooperation from anyone. Guarded by a live test that
   taskkills the worker and asserts Excel dies. The manifest sweep remains
   only for the case where job assignment itself failed.
5. Machine-wide session lock plus manifest sweep. Excel UI automation is a
   single-user surface; parallel harness sessions contend on dialogs, focus,
   and cleanup (XLIDE oracle README). Sessions take a global lock by default.
   A weakref finalizer additionally kills recorded processes when a session
   object is garbage collected or the interpreter exits without close().

## Accuracy rules

- Timeout is never VBA evidence. Only results captured inside VBA (or an
  explicit dialog capture) describe VBA behavior; a timeout means the harness
  or environment needs attention. Outcomes are literal:
  `passed | vba-error | timeout | modal-blocked | runner-error`.
- Never attach. The worker creates its own Excel with `DispatchEx` and never
  calls `GetActiveObject`; a user's open Excel is out of bounds.
- Structured trace. Every lifecycle step is an event
  (`excel-created`, `workbook-opened`, `command-started`, `modal-*`,
  `command-finished`, `workbook-closed`, `excel-quit`, `excel-killed`, ...).
  `pyvbaharness.oracle.validate_trace` checks the invariants
  (single owned instance, alerts suppressed, timeout on every command, kill
  after hang or modal-block, no commands after kill, close without save).
  Unit tests replay traces without Excel; sessions validate their own trace
  at teardown.
- Success status is read from captured results, not from exit codes. A clean
  worker exit with a missing result is a `runner-error`, not a pass.

## Measured Excel behaviors that shaped the design

Each of these was found by running the harness against Excel 365 x64 on
2026-07-25, not assumed. Each is now covered by a test.

### Dispatch attaches instead of creating

`win32com.client.Dispatch("Excel.Application")` calls `pythoncom.connect`
(GetActiveObject) before falling back to `CoCreateInstance`, so it returns
whatever Excel is already running. Two successive Dispatch calls returned the
same PID. Since the harness kills its Excel on a hang, attaching would put a
user's open workbooks in the blast radius. The worker now calls
`CoCreateInstance` directly and snapshots EXCEL.EXE PIDs before creating; if
the new Application resolves to a pre-existing process, it refuses to run
rather than proceed. Guarded by
`TestOwnership::test_sessions_never_share_an_excel_instance`.

### Direct calls, because Application.Run breaks error trapping

A VBA procedure that calls `Application.Run` and whose callee raises an error
does not catch it: the call crosses a COM boundary, so there is no VBA call
stack to unwind, and Excel shows its runtime-error dialog instead. The same
`Err.Raise 513` produced a dialog through `Application.Run` and
`caught:513|TestSource|custom failure` through a direct call.

The dispatcher therefore calls `Module.Proc` directly, which means the
generated call has to match the callee's shape exactly: `Call` for a Sub
versus assignment for a Function, and the exact argument count. `vbasig.py`
parses the declaration (handling line continuations, comments, Optional,
ParamArray, Property Get, and Declare) so arity is validated in Python before
Excel is touched; a mismatch would otherwise become a VBA compile error at
run time. Arguments are wrapped in parentheses to force ByVal, so a Variant
from the ParamArray satisfies a typed parameter. The dispatcher is cached by
(module, proc, argument count) and regenerated only when that changes, which
is why a repeat run costs 18 ms and a retarget costs 88 ms.

### Large writes wedge Excel after a macro has run

Once any macro has run in the workbook, a single `Value2` assignment covering
roughly 6000 or more cells never returns, and every Excel window reports
`IsHungAppWindow`. No dialog is involved. The boundary is not a clean cell
count: 50x100 completed in 0.26 s, 100x50 in 5.5 s, 100x64 hung, 100x100
hung. A small write beforehand avoids it, and chunked writes are both safe
and faster (10000 cells in 0.149 s as 2500-cell blocks). Reads are
unaffected (10000 cells in 0.008 s).

`write_range` therefore splits into 2000-cell chunks
(`ranges.plan_write_chunks`, unit tested for tiling without overlap or gaps).
Guarded by `TestRangesAndFiles::test_large_write_after_macro_run`, which hangs
until the watchdog fires if the chunking is removed.

### All-optional-parameter properties resolve wrongly

Under late-bound dispatch, `Range("A1").Resize(2, 2)` returns `$B$2`, not
`$A$1:$B$2`: pywin32 invokes a property whose parameters are all optional
with no arguments, and the trailing call becomes the returned object's
default `Item` member. The same applies to `Offset`. Ranges are built from
explicit corner cells via `Cells(row, column)` instead.

### Quit does not prove Excel exited

An Excel that has been made visible at any point treats itself as
user-launched and keeps running after its last automation client
disconnects. A live run ended with every harness process gone and one
responding Excel still alive; because a compile check makes Excel visible,
this leak was intermittent. `ExcelSession.close` now polls the recorded PID
after the worker exits and kills it if it outlives the cleanup grace,
recording `excel-killed` with reason `quit-did-not-terminate`. Guarded by
`TestTeardownHygiene::test_excel_dies_even_after_being_made_visible`.

### Erl is per-procedure error context

Numbering the injected source with classic BASIC line numbers is not enough
to report error lines: when an error propagates out of the erroring
procedure into the dispatcher's handler, ``Erl`` there reads the
dispatcher's own (unnumbered) lines and returns 0 (observed live,
2026-07-25). The origin line is only visible to a handler INSIDE the
erroring procedure, which is why per-procedure handler injection is the
established pattern in VBA tooling.

``numbering.instrument_error_lines`` therefore does both: numbers each
executable statement with its original physical line index, and appends a
minimal handler per procedure that calls ``PyVbaFail Erl, Err.Number,
Err.Source, Err.Description`` and re-raises. ``PyVbaFail`` records the
FIRST noted line (the deepest frame, i.e. the origin) in the support module,
and the dispatcher reads it back with ``PyVbaLastErl``. The inserted lines
carry no numbers (a numbered call line would overwrite ``Erl`` before it is
read) and declare no locals (no collisions with user names). A user ``On
Error`` statement overrides the injected one from that point on, so user
error semantics are preserved. Guarded by ``TestErrorLines``.

### The Compile control is a positive completion signal

The VBE disables its Compile command exactly when the project is fully
compiled. Polling ``control.Enabled`` after firing the compile turns the
accept verdict from "no dialog appeared during the whole watch window"
(6.1 s measured) into a positive signal that lands in about a second
(0.98 s measured). A failed compile leaves the control enabled and raises
the dialog, and the watcher's records are swept once more before an accept
is reported so a rejection can never be outrun. The watch window remains as
the fallback when neither signal appears.

### Modifying the VBProject resets VBA module-level state

Adding or replacing any module clears every module-level variable in the
project. Anything the harness pushes into the support module is silently
lost the moment a dispatcher is regenerated, which is exactly what happens
between injection and the run. Found twice: progress heartbeats stopped
arriving, and coverage hit arrays came back empty, both with no error
anywhere. The fix pattern is lazy re-push guarded by a flag that
`_write_module` clears (`_ensure_progress_path`, `_ensure_coverage`).

### PrintWindow blocks on a non-pumping window

`PrintWindow` sends a message, so on a wedged Excel (the exact case a
failure screenshot is for) it never returns. During development it wedged
the timeout path itself, converting the harness's core guarantee into a
hang. Now `IsHungAppWindow` selects `BitBlt` instead, and every capture runs
on a daemon thread with a join timeout, so no GDI call can delay a kill.

### A coverage hit cannot prefix a block opener

`PyVbaCovHit 1, 4: If flag Then` turns a block `If` into a single-line `If`,
and the matching `End If` then fails to compile, surfacing as a modal
compile-error dialog. Block openers (`If`, `For`, `Do`, `While`, `Select`,
`With`) are numbered but never prefixed, and are excluded from the coverable
set because their execution cannot be observed.

### Worksheet cells coerce staged text

Batch arguments stage through a hidden worksheet, where `"5"` would become
the number 5, `"3/4"` a date, and `"=SUM(A1)"` a formula. Every value is
type-prefixed (`s:`, `i:`, `d:`, `b:`, `e:`) so cells always hold literal
text and round-trip exactly; the VBA `DecodeArg` reverses it.

### Process liveness is not process existence

`OpenProcess` keeps succeeding for an exited process while any handle to it
remains open, so a creation-time query is identity, not liveness.
`is_process_alive` uses `GetExitCodeProcess` against `STILL_ACTIVE`; without
it, teardown checks reported a dead Excel as alive.

## Performance decisions

- Persistent session. Both references spawn Excel per run (max isolation) and
  pay 1-3 s startup each time. `ExcelSession` keeps one worker + Excel alive
  across many runs and recycles on any abnormal outcome. One-shot `run_vba`
  remains available for full isolation.
- Injection cache. `add_module` hashes (kind, transformed source); identical
  content is not resent, so `run_vba` in a loop costs a warm run instead of
  a module replacement (97 ms). The cache clears on workbook change and
  recycle, and `remove_module` evicts its entry.
- Signature cache. Resolving a run target used to read the module's source
  through VBE COM on every single run. Caching it (invalidated by any
  module write) took a warm run from 15 ms to 0.6 ms, a 25x improvement and
  the largest single performance change in the project.
- Batch execution. `run_batch` stages many calls on a hidden worksheet and
  runs them from one generated dispatcher in a single COM round trip:
  2.7x faster at 50 calls, 4.7x at 200, 9.9x at 3000 (0.066 ms per call).
  Per-call fidelity is unchanged, including error lines and stacks.
- Unsaved in-memory workbooks by default. `new_workbook()` never touches disk
  (the XLIDE syntax oracle uses the same trick to bypass file macro-security
  policy); opening an existing workbook defaults to read-only with
  `UpdateLinks:=0` and no save on close.
- Late-bound COM only. `win32com.client.dynamic` dispatch avoids the pywin32
  gencache entirely (a corrupt gencache was observed on this machine during
  environment verification). Constants are defined locally.
- Batch range IO. Reads and writes use one `Range.Value2` round trip per call
  with 2-D arrays, never per-cell COM calls.
- Cheap protocol. Newline-delimited JSON on pipes, UTF-8 forced on both ends;
  no polling loops on the supervisor side except armed timers.

## Parallel execution (SessionPool)

Parallelism is per-Excel-instance: `SessionPool` owns N full ExcelSessions,
each with its own worker process, its own new EXCEL.EXE, its own watcher,
manifest, and job object. A ThreadPoolExecutor with one thread per member
checks a session out to exactly one task at a time, which preserves the
sessions' single-caller contract without any locking inside the session.

The default machine-wide session mutex exists to stop ACCIDENTAL
concurrency; pool members run `exclusive=False` because the pool is
deliberate concurrency with the hazards addressed head-on:

- Dialogs cannot cross-talk: the watcher enumerates windows by owned PID and
  dismisses via SendMessage directly to button handles, never via focus or
  keystrokes.
- Compile checks stay serialized machine-wide through a dedicated mutex
  inside `compile_project` (pool or not): they make Excel and the VBE
  visible and drive the VBE command bar, a genuinely shared UI surface -
  the original single-user-surface warning from the XLIDE oracle README
  applies to exactly this operation.
- A hang costs one member one recycle (`auto_recycle` is forced on);
  in-flight work on other members is unaffected, proven live by wall-clock
  overlap tests.

`pool.run_tests` shards a suite by round-robin over pure-Python discovery
(`vbasig.discover_tests`), each member injects the same module and runs its
slice, and results merge back in discovery order. Session-level
`run_tests` recovery (recycle, reinject, continue) makes a hanging test
cost one test rather than one shard.

Measured scaling (16 cores, ~120 ms busy tasks, 24 tasks,
pool-baseline-0.3.0.json): 7.1 tasks/s at size 1, 12.2 at 2 (1.71x), 16.1
at 4 (2.27x), 17.0 at 6 (2.39x). The knee sits near 4 for short tasks: the
fixed ~20 ms per-run harness overhead and the supervisor's Python-side work
stop amortizing. Longer tasks scale closer to linearly. Startup is
concurrent (0.8 s for 2 members, 2.0 s for 6, warm).

## VBE object model use

Module injection uses `Workbook.VBProject.VBComponents` (`AddFromString`),
which requires the Excel option "Trust access to the VBA project object
model". This is the documented requirement of the harness. The alternative
file-staging path (writing modules into a copy with pyOpenVBA before open)
is deliberately out of scope for v1: it trades the trust setting for slower
per-change reopen cycles, and this project assumes the setting is on.

`compile_project()` invokes the VBE `Compile <project>` command-bar control
(the syntax-oracle pattern). Excel must be visible for that call: hidden
hosts do not surface the `Compile error:` dialog as a visible window, which
silently converts rejections into false accepts (XLIDE oracle lesson #1).
The session makes Excel visible for the duration of a compile check and
restores hidden mode afterward. Compile outcomes are
`accepted | rejected (dialog text) | infrastructure-failure`; a timeout is
infrastructure, never a verdict.

## Excel configuration

Applied by the worker to its owned instance:

| Setting | Value | Reason |
| --- | --- | --- |
| `Visible` | `False` (True only for compile checks) | no UI, no focus theft |
| `DisplayAlerts` | `False` | alerts become blocked automation |
| `EnableEvents` | `False` | workbook event code must not fire |
| `ScreenUpdating` | `False` | performance |
| `AskToUpdateLinks` | `False` | prompt suppression |
| `AutomationSecurity` | `msoAutomationSecurityLow (1)` | macros must run without prompts |
| `Workbooks.Open` | `UpdateLinks:=0, ReadOnly per caller, IgnoreReadOnlyRecommended:=True` | prompt suppression |
| Close | `SaveChanges:=False` unless explicitly saved | no accidental mutation |

Teardown order: stop watcher, close workbook (no save), `Quit`, release COM
references, `CoUninitialize`. Every step is timing-instrumented and covered
by the supervisor cleanup watchdog.

## Package layout

```text
src/pyvbaharness/
  __init__.py         public API (ExcelSession, SessionPool, run_vba, results)
  __main__.py         CLI: doctor / run / check
  session.py          supervisor: worker lifecycle, watchdogs, abort path
  pool.py             SessionPool: N sessions, work queue, sharded run_tests
  protocol.py         command/event types and JSON codec
  results.py          RunResult, CompileResult, coverage, outcomes
  codegen.py          support, dispatcher, and batch module generation
  vbasig.py           VBA declaration parsing (kind, arity, param types)
  numbering.py        line numbering, Erl handlers, coverage hooks
  ranges.py           write chunk planning and block validation
  dialog_policy.py    pure dialog classification / dismissal decision
  oracle.py           pure trace validator
  process_control.py  taskkill, manifests, stale sweep, kill-on-close job
  lock.py             session and compile mutexes
  screenshot.py       bounded GDI window capture + minimal PNG encoder
  properties.py       Hypothesis strategies from VBA signatures
  pytest_plugin.py    .bas collection as pytest items
  worker/__main__.py  worker entry point, command loop, progress tail
  worker/excel_host.py  all Excel COM calls
  worker/watcher.py   ctypes window scanner + dismissal executor
tests/unit            pure logic, no Excel required (147 tests)
tests/live            real Excel, opt-in via -m live (54 tests)
benchmarks/           per-run, batch, and pool scaling measurements
```

How to change this codebase, including the full catalog of measured Excel
behaviors and the recipe for adding a capability end to end, is in
[IMPLEMENTATION_GUIDE.md](IMPLEMENTATION_GUIDE.md).

Pure-logic modules (`dialog_policy`, `oracle`, `codegen`, `vbasig`, `ranges`,
`protocol`) hold everything unit-testable so the live suite stays small and
targeted. `tests/unit/fake_worker.py` speaks the pipe protocol with scripted
hangs, modal blocks, and crashes, so the supervisor's watchdog, abort, and
recycle paths are tested without Excel.

## Diagnostics

Setting `PYVBAHARNESS_STACK_DUMP_S=<seconds>` makes the worker dump every
thread's stack to stderr on that interval. The supervisor tails stderr, so a
wedged command can be located from the outside. This is how the large-write
wedge above was traced to the `Value2` assignment rather than guessed at.

## Known limits (v1)

- Windows desktop Excel only (win32, pywin32).
- Break-mode code-line capture (ROneCOne's `GetSelection` trick) is not
  implemented: the watcher thread cannot call COM into the blocked STA. A
  break is detected via the VBE window and handled as blocked; the offending
  error text usually arrives via dialog capture instead.
- The dispatcher supports 10 arguments, which is checked and reported rather
  than silently truncated.
- A target returning an object fails at the dispatcher's Variant assignment
  and is reported as a vba-error. Return scalars or arrays, or write to
  cells.
- One session per machine by default (lock); deliberate parallelism goes
  through SessionPool, which opts members out and keeps compile checks
  serialized. Concurrent calls on one session object remain out of contract
  (the pool enforces one task per member).
- No warm-spare Excel pool. A pre-spawned spare would make recycle-after-hang
  nearly instant (0.5-3 s today), but its event stream would interleave into
  the session trace mid-generation and it complicates the abort path, the
  most safety-critical code in the harness. SessionPool reduces the pain a
  different way: while one member recycles, the others keep serving.
  Declined until a real workload shows recycle latency itself to be the
  bottleneck; the design note is here so the tradeoff is not re-derived
  from scratch.
