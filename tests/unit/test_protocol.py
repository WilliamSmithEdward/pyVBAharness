import pytest

from pyvbaharness import protocol


class TestCommands:
    def test_roundtrip(self):
        line = protocol.encode_command(7, "run", {"target": "M.P",
                                                  "args": [1, "x"]})
        decoded = protocol.decode_command(line)
        assert decoded["cid"] == 7
        assert decoded["cmd"] == "run"
        assert decoded["params"] == {"target": "M.P", "args": [1, "x"]}

    def test_unicode_survives(self):
        line = protocol.encode_command(1, "run", {"target": "Café"})
        assert protocol.decode_command(line)["params"]["target"] == "Café"

    def test_malformed_rejected(self):
        for bad in ('{"cmd": "x"}', '{"cid": 1}', '[]',
                    '{"cid": 1, "cmd": "x", "params": []}'):
            with pytest.raises(ValueError):
                protocol.decode_command(bad)


class TestEvents:
    def test_roundtrip(self):
        line = protocol.encode_event({"kind": "worker-ready", "pid": 5})
        assert line.startswith(protocol.EVENT_PREFIX)
        event = protocol.decode_event(line)
        assert event == {"kind": "worker-ready", "pid": 5}

    def test_plain_lines_pass_through(self):
        assert protocol.decode_event("hello world") is None

    def test_structurally_wrong_json_rejected(self):
        assert protocol.decode_event(protocol.EVENT_PREFIX + '["x"]') is None
        assert protocol.decode_event(protocol.EVENT_PREFIX + '{"a": 1}') is None

    def test_unserializable_degrades_to_string(self):
        class Odd:
            def __str__(self) -> str:
                return "odd-thing"

        line = protocol.encode_event({"kind": "k", "value": Odd()})
        assert protocol.decode_event(line)["value"] == "odd-thing"
