import json
import os
import time

from pyvbaharness.process_control import (
    OwnedProcessManifest,
    pid_matches,
    process_creation_time,
    sweep_stale_manifests,
)


class TestCreationTime:
    def test_own_process_queryable(self):
        stamp = process_creation_time(os.getpid())
        assert isinstance(stamp, int) and stamp > 0

    def test_dead_pid_unqueryable(self):
        # PID 4 is System (unqueryable at limited rights is fine too); use an
        # absurd PID that cannot exist instead.
        assert process_creation_time(0x7FFFFFF0) is None

    def test_pid_matches_guards_reuse(self):
        pid = os.getpid()
        real = process_creation_time(pid)
        assert pid_matches(pid, real)
        assert pid_matches(pid, None)  # no recorded stamp: existence check
        assert not pid_matches(pid, real + 1)


class TestManifest:
    def test_record_and_entry(self, tmp_path):
        manifest = OwnedProcessManifest("s1", tmp_path)
        manifest.record("worker", os.getpid())
        pid, creation = manifest.entry("worker")
        assert pid == os.getpid()
        assert creation == process_creation_time(os.getpid())
        on_disk = json.loads(manifest.path.read_text(encoding="utf-8"))
        assert on_disk["worker"]["pid"] == os.getpid()
        manifest.remove()
        assert not manifest.path.exists()

    def test_missing_role(self, tmp_path):
        manifest = OwnedProcessManifest("s2", tmp_path)
        assert manifest.entry("excel") is None
        assert manifest.kill_role("excel") is False


class TestSweep:
    def test_fresh_manifest_untouched(self, tmp_path):
        manifest = OwnedProcessManifest("fresh", tmp_path)
        manifest.record("worker", os.getpid())
        notes = sweep_stale_manifests(tmp_path)
        assert notes == []
        assert manifest.path.exists()

    def test_stale_manifest_with_dead_pids_deleted(self, tmp_path):
        path = tmp_path / "old.json"
        path.write_text(json.dumps({
            "session": "old",
            "written_at": time.time() - 10_000,
            "excel": {"pid": 0x7FFFFFF0, "creation": 1},
        }), encoding="utf-8")
        notes = sweep_stale_manifests(tmp_path, stale_after_s=60)
        assert notes == []  # nothing alive to kill
        assert not path.exists()  # stale bookkeeping removed
