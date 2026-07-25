"""Measure SessionPool scaling against live Excel.

    python benchmarks/run_pool_benchmarks.py [--out benchmarks/output/pool.json]

Runs a fixed batch of CPU-holding VBA tasks (each ~120 ms of busy work in
its Excel instance) through pools of increasing size and reports wall time,
throughput, and speedup versus size 1. Startup cost is reported separately
so the steady-state scaling is visible. The knee is where speedup stops
tracking pool size; expect it near the physical core count, minus whatever
the machine is doing besides Excel.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyvbaharness import SessionPool, __version__  # noqa: E402

TASKS = 24
BUSY_SECONDS = 0.12

BUSY_SOURCE = f"""
Public Sub Busy()
    Dim started As Single
    started = Timer
    Do While Timer - started < {BUSY_SECONDS}!
    Loop
End Sub
"""


def measure(size: int) -> dict:
    created = time.perf_counter()
    with SessionPool(size) as pool:
        startup_s = time.perf_counter() - created
        # Warm every member (inject + first dispatch) before timing.
        warm = [pool.run_vba(BUSY_SOURCE, proc="Busy", timeout=30.0)
                for _ in range(size)]
        for future in warm:
            future.result(timeout=60)
        started = time.perf_counter()
        futures = [pool.run_vba(BUSY_SOURCE, proc="Busy", timeout=30.0)
                   for _ in range(TASKS)]
        outcomes = [f.result(timeout=120).outcome for f in futures]
        wall_s = time.perf_counter() - started
    failed = sum(1 for outcome in outcomes if outcome != "passed")
    return {
        "pool_size": size,
        "startup_s": round(startup_s, 3),
        "tasks": TASKS,
        "wall_s": round(wall_s, 3),
        "tasks_per_s": round(TASKS / wall_s, 2),
        "failed": failed,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default="")
    parser.add_argument("--sizes", default="1,2,4,6")
    args = parser.parse_args()
    sizes = [int(s) for s in args.sizes.split(",") if s.strip()]

    rows = [measure(size) for size in sizes]
    baseline = rows[0]["wall_s"]
    for row in rows:
        row["speedup"] = round(baseline / row["wall_s"], 2)

    payload = {
        "version": __version__,
        "python": sys.version.split()[0],
        "cpu_count": os.cpu_count(),
        "task_busy_seconds": BUSY_SECONDS,
        "results": rows,
    }
    text = json.dumps(payload, indent=2)
    print(text)
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text + "\n", encoding="utf-8")
    return 1 if any(row["failed"] for row in rows) else 0


if __name__ == "__main__":
    sys.exit(main())
