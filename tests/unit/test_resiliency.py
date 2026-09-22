"""Scoped removal of Office crash-recovery entries.

These run against a private registry tree rather than the real one, so the
matching and, more importantly, the refusal to touch anything else can be
tested without depending on what Office happens to have pending.
"""
import winreg

import pytest

from pyvbaharness import apps, resiliency

OURS = r"C:\work\model.xlsm"
THEIRS = r"F:\someone-else\fixtures\DebugFixture.xlsm"


def _blob(path: str) -> bytes:
    """Stand-in for the recovery blob: undocumented binary carrying the
    document path as UTF-16, which is all the matcher reads."""
    return b"\x01\x00\xff\xfe" + path.encode("utf-16-le") + b"\x00\x00"


@pytest.fixture
def fake_tree(tmp_path, monkeypatch):
    """A DocumentRecovery tree under HKCU\\Software\\pyvbaharness-test."""
    root = r"Software\pyvbaharness-test"
    created: list[str] = []

    def make(app_key: str, entries: dict[str, str]) -> None:
        base = (f"{root}\\16.0\\{apps.info(app_key).registry_key}"
                f"\\Resiliency\\DocumentRecovery")
        for name, path in entries.items():
            key_path = f"{base}\\{name}"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
            winreg.SetValueEx(key, name, 0, winreg.REG_BINARY, _blob(path))
            winreg.CloseKey(key)
            created.append(key_path)

    monkeypatch.setattr(resiliency, "_OFFICE_ROOT", root)
    monkeypatch.setattr(resiliency, "office_versions", lambda: ["16.0"])
    yield make

    def drop(path: str) -> None:
        for child in resiliency._subkeys(winreg.HKEY_CURRENT_USER, path):
            drop(f"{path}\\{child}")
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except OSError:
            pass

    drop(root)


def _remaining(app_key: str) -> set[str]:
    base = (r"Software\pyvbaharness-test\16.0"
            f"\\{apps.info(app_key).registry_key}"
            r"\Resiliency\DocumentRecovery")
    return set(resiliency._subkeys(winreg.HKEY_CURRENT_USER, base))


class TestScoping:
    def test_removes_an_entry_naming_our_document(self, fake_tree):
        fake_tree("excel", {"AAA111": OURS})
        removed = resiliency.clear_document_recovery("excel", [OURS])
        assert [e.key_path.rsplit("\\", 1)[-1] for e in removed] == ["AAA111"]
        assert _remaining("excel") == set()

    def test_leaves_entries_belonging_to_anything_else(self, fake_tree):
        """The case that makes a blanket wipe unacceptable: the same tree
        holds pending recovery data the harness never created."""
        fake_tree("excel", {"AAA111": OURS, "BBB222": THEIRS})
        removed = resiliency.clear_document_recovery("excel", [OURS])
        assert len(removed) == 1
        assert _remaining("excel") == {"BBB222"}

    def test_no_paths_removes_nothing(self, fake_tree):
        fake_tree("excel", {"BBB222": THEIRS})
        assert resiliency.clear_document_recovery("excel", []) == []
        assert _remaining("excel") == {"BBB222"}

    def test_empty_path_strings_are_ignored(self, fake_tree):
        """An unsaved document reports no path, and an empty string would
        otherwise substring-match every entry there is."""
        fake_tree("excel", {"BBB222": THEIRS})
        assert resiliency.clear_document_recovery("excel", ["", ""]) == []
        assert _remaining("excel") == {"BBB222"}

    def test_matching_ignores_case(self, fake_tree):
        fake_tree("excel", {"AAA111": OURS})
        removed = resiliency.clear_document_recovery("excel", [OURS.upper()])
        assert len(removed) == 1

    def test_only_the_named_app_is_touched(self, fake_tree):
        fake_tree("excel", {"AAA111": OURS})
        fake_tree("word", {"CCC333": OURS})
        resiliency.clear_document_recovery("excel", [OURS])
        assert _remaining("excel") == set()
        assert _remaining("word") == {"CCC333"}

    def test_a_missing_tree_is_not_an_error(self, fake_tree):
        assert resiliency.clear_document_recovery("access", [OURS]) == []


class TestVersionDiscovery:
    def test_versions_are_enumerated_not_assumed(self):
        """A machine can have more than one Office version, and the
        fallback is only used when none are present."""
        versions = resiliency.office_versions()
        assert versions
        assert all(v.replace(".", "").isdigit() for v in versions)
        assert versions == sorted(versions, key=float, reverse=True)
