"""The fuzz corpus, replayed through the fuzz targets on every run.

tests/fuzz_corpus/<target> seeds fuzz/fuzz_harness.py; a fuzz finding joins
it as a regression seed.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from pyvbaharness import protocol, vbasig

CORPUS = Path(__file__).resolve().parent.parent / "fuzz_corpus"
VBASIG = sorted((CORPUS / "vbasig").iterdir())
PROTOCOL = sorted((CORPUS / "protocol").iterdir())


def _text(path: Path) -> str:
    return path.read_bytes().decode("utf-8", errors="replace")


@pytest.mark.parametrize("path", VBASIG, ids=lambda p: p.name)
def test_a_module_parses_consistently(path: Path) -> None:
    source = _text(path)
    for signature in vbasig.list_procedures(source):
        found = vbasig.find_procedure(source, signature.name)
        assert found is not None and found.name.lower() == signature.name.lower()
        assert signature.accepts(signature.required)
    callable_now = {s.name for s in vbasig.list_procedures(source)
                    if s.kind in (vbasig.KIND_SUB, vbasig.KIND_FUNCTION) and s.required == 0}
    assert set(vbasig.discover_tests(source)) <= callable_now


@pytest.mark.parametrize("path", PROTOCOL, ids=lambda p: p.name)
def test_a_line_decodes_or_is_refused_as_garbled(path: Path) -> None:
    line = _text(path)
    try:
        event = protocol.decode_event(line)
    except ValueError:
        event = None
    assert event is None or isinstance(event["kind"], str)
    try:
        command = protocol.decode_command(line)
    except ValueError:
        return
    assert isinstance(command["cmd"], str) and isinstance(command["cid"], int)
    assert isinstance(command["params"], dict)
