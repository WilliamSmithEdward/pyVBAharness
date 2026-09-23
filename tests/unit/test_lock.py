"""Machine-wide mutexes, and the part of exclusivity the kernel cannot see.

A Windows mutex is owned by the thread that waited on it and is recursive,
so a second wait on the same thread returns immediately. Measured
2026-09-22: two exclusive sessions built in one thread both believed they
held the session lock, and neither was told otherwise.
"""
import threading
import time

import pytest

from pyvbaharness.lock import (
    COMPILE_MUTEX_NAME,
    CREATE_MUTEX_NAME,
    SESSION_MUTEX_NAME,
    SessionLock,
    session_mutex_name,
)
from pyvbaharness.results import SessionLockHeld

NAME = "Global\\pyvbaharness-unit-test-lock"
OTHER = "Global\\pyvbaharness-unit-test-lock-other"


class TestInProcessExclusivity:
    def test_a_second_holder_in_one_thread_is_refused(self):
        first = SessionLock(timeout_s=0.0, name=NAME)
        second = SessionLock(timeout_s=0.0, name=NAME)
        first.acquire()
        try:
            with pytest.raises(SessionLockHeld):
                second.acquire()
        finally:
            first.release()

    def test_a_second_holder_on_another_thread_is_refused(self):
        held = SessionLock(timeout_s=0.0, name=NAME)
        held.acquire()
        outcome = []

        def contender() -> None:
            try:
                SessionLock(timeout_s=0.0, name=NAME).acquire()
                outcome.append("acquired")
            except SessionLockHeld:
                outcome.append("refused")

        thread = threading.Thread(target=contender)
        thread.start()
        thread.join(timeout=10)
        held.release()
        assert outcome == ["refused"]

    def test_release_lets_the_next_holder_in(self):
        first = SessionLock(timeout_s=0.0, name=NAME)
        first.acquire()
        first.release()
        second = SessionLock(timeout_s=0.0, name=NAME)
        second.acquire()
        second.release()

    def test_a_refused_acquire_does_not_consume_the_name(self):
        """A failed attempt must not leave the name marked as held, or the
        real holder's release would free nothing."""
        first = SessionLock(timeout_s=0.0, name=NAME)
        first.acquire()
        with pytest.raises(SessionLockHeld):
            SessionLock(timeout_s=0.0, name=NAME).acquire()
        first.release()
        again = SessionLock(timeout_s=0.0, name=NAME)
        again.acquire()
        again.release()

    def test_different_names_are_independent(self):
        one = SessionLock(timeout_s=0.0, name=NAME)
        two = SessionLock(timeout_s=0.0, name=OTHER)
        one.acquire()
        two.acquire()
        two.release()
        one.release()

    def test_release_without_acquire_is_harmless(self):
        SessionLock(timeout_s=0.0, name=NAME).release()
        taken = SessionLock(timeout_s=0.0, name=NAME)
        taken.acquire()
        taken.release()

    def test_the_context_manager_releases(self):
        with SessionLock(timeout_s=0.0, name=NAME):
            pass
        after = SessionLock(timeout_s=0.0, name=NAME)
        after.acquire()
        after.release()


class TestWaitingRatherThanRefusing:
    def test_a_waiting_holder_gets_in_when_the_first_releases(self):
        """Compile checks inside a pool queue on this lock. Refusing on
        contention instead of waiting would fail a sibling's compile rather
        than serialize it."""
        first = SessionLock(timeout_s=0.0, name=NAME, purpose="compile")
        first.acquire()
        got = []

        def waiter() -> None:
            lock = SessionLock(timeout_s=10.0, name=NAME, purpose="compile")
            try:
                lock.acquire()
                got.append("acquired")
                lock.release()
            except SessionLockHeld:
                got.append("refused")

        thread = threading.Thread(target=waiter)
        thread.start()
        time.sleep(0.3)
        assert got == [], "the waiter should still be queued"
        first.release()
        thread.join(timeout=15)
        assert got == ["acquired"]


class TestNames:
    def test_excel_keeps_the_original_name(self):
        """A session started by 1.x still excludes one started by a newer
        version."""
        assert session_mutex_name("excel") == SESSION_MUTEX_NAME

    @pytest.mark.parametrize("app", ["word", "powerpoint", "access"])
    def test_other_apps_get_their_own(self, app):
        name = session_mutex_name(app)
        assert name != SESSION_MUTEX_NAME
        assert app in name

    def test_the_three_purposes_are_distinct(self):
        assert len({SESSION_MUTEX_NAME, COMPILE_MUTEX_NAME,
                    CREATE_MUTEX_NAME}) == 3


class TestMessages:
    @pytest.mark.parametrize("purpose, expected", [
        ("session", "SessionPool"),
        ("compile", "compile check"),
        ("create", "Office instance"),
    ])
    def test_each_purpose_explains_itself(self, purpose, expected):
        held = SessionLock(timeout_s=0.0, name=NAME, purpose=purpose)
        held.acquire()
        try:
            with pytest.raises(SessionLockHeld, match=expected):
                SessionLock(timeout_s=0.0, name=NAME,
                            purpose=purpose).acquire()
        finally:
            held.release()
