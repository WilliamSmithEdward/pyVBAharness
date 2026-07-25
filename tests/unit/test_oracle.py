from pyvbaharness import protocol as p
from pyvbaharness.oracle import validate_trace


def created(**overrides):
    event = {"kind": p.EV_EXCEL_CREATED, "pid": 123, "attached": False,
             "display_alerts": False, "enable_events": False}
    event.update(overrides)
    return event


def started(cid=1, timeout_ms=30000, command="run"):
    return {"kind": p.EV_COMMAND_STARTED, "cid": cid, "command": command,
            "timeout_ms": timeout_ms}


def finished(cid=1, outcome="passed", command="run"):
    return {"kind": p.EV_COMMAND_FINISHED, "cid": cid, "command": command,
            "outcome": outcome}


def happy_trace():
    return [
        created(),
        {"kind": p.EV_READY, "pid": 123},
        {"kind": p.EV_WORKBOOK_CREATED, "name": "Book1"},
        started(),
        finished(),
        {"kind": p.EV_WORKBOOK_CLOSED, "save_changes": False},
        {"kind": p.EV_EXCEL_QUIT},
    ]


class TestOracle:
    def test_happy_trace_is_clean(self):
        assert validate_trace(happy_trace()) == []

    def test_empty_trace(self):
        issues = validate_trace([])
        assert [i.code for i in issues] == ["empty-trace"]

    def test_attached_instance_flagged(self):
        issues = validate_trace([created(attached=True),
                                 {"kind": p.EV_EXCEL_QUIT}])
        assert "attached-excel-instance" in [i.code for i in issues]

    def test_alerts_must_be_suppressed(self):
        issues = validate_trace([created(display_alerts=True),
                                 {"kind": p.EV_EXCEL_QUIT}])
        assert "suppress-alerts" in [i.code for i in issues]

    def test_missing_timeout_flagged(self):
        trace = [created(), started(timeout_ms=None), finished(),
                 {"kind": p.EV_EXCEL_QUIT}]
        assert "command-timeout" in [i.code for i in validate_trace(trace)]

    def test_hang_requires_kill(self):
        trace = [created(), started(), finished(outcome="timeout")]
        assert "hang-cleanup" in [i.code for i in validate_trace(trace)]

    def test_hang_followed_by_kill_is_clean(self):
        trace = [created(), started(), finished(outcome="timeout"),
                 {"kind": p.EV_EXCEL_KILLED, "reason": "timeout"}]
        assert validate_trace(trace) == []

    def test_no_commands_after_kill(self):
        trace = [created(),
                 {"kind": p.EV_EXCEL_KILLED, "reason": "timeout"},
                 started(cid=2)]
        assert "no-commands-after-kill" in [i.code
                                            for i in validate_trace(trace)]

    def test_close_with_save_flagged_without_save_command(self):
        trace = [created(),
                 {"kind": p.EV_WORKBOOK_CLOSED, "save_changes": True},
                 {"kind": p.EV_EXCEL_QUIT}]
        assert "close-without-saving" in [i.code
                                          for i in validate_trace(trace)]

    def test_close_with_save_allowed_after_save_command(self):
        trace = [created(),
                 started(command=p.CMD_SAVE_AS),
                 finished(command=p.CMD_SAVE_AS),
                 {"kind": p.EV_WORKBOOK_CLOSED, "save_changes": True},
                 {"kind": p.EV_EXCEL_QUIT}]
        assert validate_trace(trace) == []

    def test_session_must_end_excel(self):
        trace = [created(), started(), finished()]
        assert "normal-cleanup" in [i.code for i in validate_trace(trace)]

    def test_two_instances_flagged(self):
        trace = [created(), created(), {"kind": p.EV_EXCEL_QUIT}]
        assert "single-owned-excel-instance" in [
            i.code for i in validate_trace(trace)]
