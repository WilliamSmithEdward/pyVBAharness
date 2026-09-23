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

    def test_a_long_running_session_is_spared(self, tmp_path):
        """A manifest goes stale on the clock, not on death. A session that
        outlives the threshold is still live and its pids still match, so
        the age test alone would kill what the sweep exists to protect.
        """
        path = tmp_path / "longrun.json"
        path.write_text(json.dumps({
            "session": "longrun",
            "written_at": time.time() - 10_000,
            # This process stands in for a worker that is still running.
            "worker": {"pid": os.getpid(),
                       "creation": process_creation_time(os.getpid())},
            "app": {"pid": os.getpid(),
                    "creation": process_creation_time(os.getpid())},
        }), encoding="utf-8")
        notes = sweep_stale_manifests(tmp_path, stale_after_s=60)
        assert notes == []
        assert path.exists(), "a live session's manifest was swept"

    def test_a_dead_worker_still_reaps_its_host(self, tmp_path):
        """The case the sweep exists for: the worker has gone and the host
        outlived it. A pid that cannot match anything stands in for both."""
        path = tmp_path / "orphan.json"
        path.write_text(json.dumps({
            "session": "orphan",
            "written_at": time.time() - 10_000,
            "worker": {"pid": 0x7FFFFFF0, "creation": 1},
            "app": {"pid": 0x7FFFFFF1, "creation": 1},
        }), encoding="utf-8")
        sweep_stale_manifests(tmp_path, stale_after_s=60)
        assert not path.exists()

    def test_a_manifest_without_a_worker_is_not_assumed_live(self, tmp_path):
        """Older manifests, and any written before the worker was recorded,
        cannot prove liveness and stay sweepable."""
        path = tmp_path / "noworker.json"
        path.write_text(json.dumps({
            "session": "noworker",
            "written_at": time.time() - 10_000,
            "app": {"pid": 0x7FFFFFF0, "creation": 1},
        }), encoding="utf-8")
        sweep_stale_manifests(tmp_path, stale_after_s=60)
        assert not path.exists()
