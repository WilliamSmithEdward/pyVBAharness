# Agent Instructions

These instructions apply to the entire repository.

## Governing models

Use the Recursive Invariant Discovery Model (RIDM) 11.0 as the operating
model, and AI Best Practices as the engineering and writing baseline:

- https://github.com/WilliamSmithEdward/RIDM_Recursive_Invariant_Discovery_Model
- https://github.com/WilliamSmithEdward/AI_Best_Practices

Keep both mostly invisible in user-facing wording.

## Mission

pyVBAharness runs VBA inside real desktop Excel without hanging the calling
process. When goals conflict, the priority is hang resistance, then accuracy,
then performance. Make the smallest coherent change that preserves that
order.

## Project specifics

- Toolchain: Python 3.10+, pywin32, Windows desktop Excel.
- Unit tests (no Excel): `python -m pytest tests/unit`
- Live tests (real Excel): `python -m pytest tests/live -m live -o addopts=""`
- Benchmarks: `python benchmarks/run_benchmarks.py --out benchmarks/output/<name>.json`
- Architecture and the measured Excel behaviors behind each decision:
  `docs/architecture.md`.
- Excel's "Trust access to the VBA project object model" must be enabled.

## Invariants

Do not weaken these without changing the documented contract in the same
change:

- The harness creates its own Excel instance and never attaches to a running
  one. `ExcelHost.create` proves this with a PID snapshot and refuses
  otherwise.
- The owned Excel lives inside a kill-on-close kernel job tied to the worker
  process, so worker death of any kind reaps Excel. Do not remove the job
  without replacing the crash-safety guarantee it provides.
- Every command carries a positive timeout, and a breach kills the recorded
  Excel process, kills the worker, and marks the session dead.
- No command runs after the owned Excel is killed.
- A timeout or blocked modal is infrastructure state, never evidence about
  the VBA under test. Only `passed` and `vba-error` describe the code.
- VBA errors are captured inside VBA by a directly-called dispatcher. Do not
  route runs through `Application.Run` from inside VBA: error trapping breaks.
- Error lines come from per-procedure Erl instrumentation (numbering.py).
  `Erl` is per-procedure error context; reading it only in the dispatcher
  returns 0, so the injected in-procedure handlers are load-bearing.
- Dialogs offering a real choice are never dismissed by guessing.
- All COM lives in the worker process; the supervisor and the dialog watcher
  stay COM-free.
- Workbooks open read-only by default and close without saving unless an
  explicit save ran.

`oracle.py` encodes most of these as a trace validator. Sessions validate
their own trace on close, and the unit suite replays synthetic traces, so a
regression in safety behavior fails loudly.

## Working on Excel behavior

Excel's automation surface has sharp edges that are not documented anywhere
useful. When one is suspected:

1. Reproduce it in a standalone probe before changing harness code.
2. Bisect until the trigger is isolated, rather than fixing by hypothesis.
3. Use `PYVBAHARNESS_STACK_DUMP_S` to locate a wedge instead of guessing.
4. Record the measurement (numbers, date, Excel build) in a code comment at
   the workaround and in `docs/architecture.md`.
5. Add a live test that fails without the workaround.

Do not treat a passing run as proof that a workaround is unnecessary; several
of the behaviors here only appear in specific orders (a large write only
wedges after a macro has run, and only when no smaller write preceded it).

## Change discipline

- Read the relevant module and its tests before editing.
- Keep pure logic in pure modules so it stays unit testable without Excel.
- Run the unit suite on every change and the live suite whenever COM,
  codegen, dialog policy, or process control changes.
- Verify no Excel processes and no manifests survive a live run:
  `Get-Process EXCEL` and `%LOCALAPPDATA%\pyvbaharness\sessions`.
- Do not commit, push, or publish unless asked.

## Writing style

Plain ASCII. No em dashes, curly quotes, or decorative emoji. Direct, specific
prose; state the measurement rather than the impression. Comments explain
constraints the code cannot show, especially the Excel behaviors above.

## Final report

Report what changed and why, the files affected, the validation actually run
with its observed results, and any remaining risk or skipped check. Do not
report completion while a material success criterion is unverified.
