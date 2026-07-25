"""Measure harness overhead against live Excel.

    python benchmarks/run_benchmarks.py [--out benchmarks/output/latest.json]

Reports medians so a single cold outlier does not set the number. What each
measurement covers:

  session_startup      ExcelSession() -> ready (new Excel process + watcher)
  first_run            inject support+dispatcher modules, then run
  warm_run             same target again (dispatcher cached: pure call cost)
  run_vba_cached       repeat run_vba with identical source (injection cache)
  retarget_run         different target (dispatcher regenerated)
  run_with_args        two arguments marshaled through the dispatcher
  write_10k / read_10k one COM round trip for 10000 cells
  compile_accept       VBE compile check (positive fast-accept signal)
"""
from __future__ import annotations

import argparse
import json
import statistics
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pyvbaharness import ExcelSession, HarnessConfig  # noqa: E402
from pyvbaharness import __version__  # noqa: E402

SOURCE_A = """
Public Function Noop() As Long
    Noop = 1
End Function

Public Function AddTwo(ByVal a As Long, ByVal b As Long) As Long
    AddTwo = a + b
End Function
"""

SOURCE_B = """
Public Function Other() As Long
    Other = 2
End Function
"""


def timed(func, repeat: int) -> dict[str, float]:
    samples = []
    for _ in range(repeat):
        started = time.perf_counter()
        func()
        samples.append(time.perf_counter() - started)
    return {
        "median_s": round(statistics.median(samples), 4),
        "min_s": round(min(samples), 4),
        "max_s": round(max(samples), 4),
        "samples": len(samples),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default="")
    parser.add_argument("--repeat", type=int, default=5)
    args = parser.parse_args()

    results: dict[str, dict] = {}

    def start_and_close():
        session = ExcelSession(HarnessConfig(exclusive=False))
        session.close()

    results["session_startup"] = timed(start_and_close, 3)

    session = ExcelSession(HarnessConfig(exclusive=False))
    try:
        session.new_workbook()
        session.add_module("BenchA", SOURCE_A)
        session.add_module("BenchB", SOURCE_B)

        first = {"done": False}

        def first_run():
            if first["done"]:
                return
            session.run_macro("BenchA.Noop")
            first["done"] = True

        started = time.perf_counter()
        first_run()
        results["first_run"] = {"median_s": round(time.perf_counter() - started, 4),
                                "samples": 1}

        results["warm_run"] = timed(
            lambda: session.run_macro("BenchA.Noop"), args.repeat)
        results["run_vba_cached"] = timed(
            lambda: session.run_vba(SOURCE_A, proc="Noop"), args.repeat)
        results["run_with_args"] = timed(
            lambda: session.run_macro("BenchA.AddTwo", 2, 3), args.repeat)

        toggle = {"flag": False}

        def retarget():
            toggle["flag"] = not toggle["flag"]
            target = "BenchA.Noop" if toggle["flag"] else "BenchB.Other"
            session.run_macro(target)

        results["retarget_run"] = timed(retarget, args.repeat)

        block = [[float(r * 100 + c) for c in range(100)] for r in range(100)]
        results["write_10k_cells"] = timed(
            lambda: session.write_range("Sheet1", "A1", block), 3)
        results["read_10k_cells"] = timed(
            lambda: session.read_range("Sheet1", "A1:CV100"), 3)

        results["compile_accept"] = timed(
            lambda: session.compile_project(watch_seconds=5.0), 1)
    finally:
        session.close()

    payload = {
        "version": __version__,
        "python": sys.version.split()[0],
        "results": results,
    }
    text = json.dumps(payload, indent=2)
    print(text)
    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
