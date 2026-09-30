"""Coverage-guided fuzzing of the parsers that read text the harness does not control.

  vbasig    the procedure-signature parser, on arbitrary VBA source. It
            never raises; every procedure list_procedures() reports is what
            find_procedure() finds by that name, and accepts its own
            required argument count; discover_tests() names only listed
            Subs and Functions that take no required argument.
  protocol  the worker protocol decoders, on an arbitrary line. A garbled
            line raises ValueError, which the session and the worker catch;
            anything else escapes them. A decoded event has a text kind, a
            decoded command a text name, an integer id and a params dict.

The package imports Windows modules on import (winreg, user32 through
ctypes), and Atheris ships Linux wheels, so the two modules, which import
only the standard library, are loaded from their files and instrumented.

    python fuzz/fuzz_harness.py <target> [libFuzzer options] [corpus dirs]
    python fuzz/fuzz_harness.py vbasig -max_total_time=60 tests/fuzz_corpus/vbasig

The .github/workflows/fuzz.yml workflow runs each target from
tests/fuzz_corpus/<target>. A finding becomes a seed there, which
tests/unit/test_fuzz_corpus.py replays on every CI run.
"""

import importlib.util
import sys
from pathlib import Path

import atheris

SRC = Path(__file__).resolve().parent.parent / "src" / "pyvbaharness"


def _load(name):
    spec = importlib.util.spec_from_file_location(f"pyvbaharness_{name}", SRC / f"{name}.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


vbasig = _load("vbasig")
protocol = _load("protocol")
atheris.instrument_all()


def fuzz_vbasig(data):
    source = data.decode("utf-8", errors="replace")
    for signature in vbasig.list_procedures(source):
        found = vbasig.find_procedure(source, signature.name)
        if found is None or found.name.lower() != signature.name.lower():
            raise AssertionError(f"find_procedure does not find {signature.name!r}")
        if not signature.accepts(signature.required):
            raise AssertionError(f"{signature.name!r} does not accept its required count")
    # Checked against the listed signatures, not find_procedure: a name
    # declared twice is a compile error in VBA, and either may come first.
    callable_now = {s.name for s in vbasig.list_procedures(source)
                    if s.kind in (vbasig.KIND_SUB, vbasig.KIND_FUNCTION) and s.required == 0}
    for name in vbasig.discover_tests(source):
        if name not in callable_now:
            raise AssertionError(f"discover_tests named {name!r}, which is not a no-argument procedure")


def fuzz_protocol(data):
    line = data.decode("utf-8", errors="replace")
    try:
        event = protocol.decode_event(line)
    except ValueError:
        event = None
    if event is not None and not isinstance(event.get("kind"), str):
        raise AssertionError("an event without a text kind was accepted")
    try:
        command = protocol.decode_command(line)
    except ValueError:
        return
    if not (isinstance(command["cmd"], str) and isinstance(command["cid"], int)
            and isinstance(command["params"], dict)):
        raise AssertionError("a malformed command was accepted")


TARGETS = {"vbasig": fuzz_vbasig, "protocol": fuzz_protocol}


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"usage: fuzz_harness.py <{'|'.join(TARGETS)}> [libFuzzer options]")
    atheris.Setup([sys.argv[0], *sys.argv[2:]], TARGETS[sys.argv[1]])
    atheris.Fuzz()


if __name__ == "__main__":
    main()
