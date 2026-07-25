"""Packaging invariants.

The version appears in two files, and a release that ships them out of sync
publishes a wheel whose reported __version__ disagrees with its filename.
Benchmark baselines stamp __version__ into their JSON, so the drift would
be recorded in results too. Cheap to check here, expensive to notice later.
"""
from pathlib import Path

import pyvbaharness

try:  # tomllib is 3.11+; the package supports 3.10, so fall back.
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - version dependent
    import tomli as tomllib

ROOT = Path(__file__).resolve().parents[2]
PYPROJECT = ROOT / "pyproject.toml"


def project_table() -> dict:
    with PYPROJECT.open("rb") as handle:
        return tomllib.load(handle)["project"]


class TestVersion:
    def test_matches_pyproject(self):
        assert pyvbaharness.__version__ == project_table()["version"]

    def test_is_pep440_release(self):
        parts = pyvbaharness.__version__.split(".")
        assert len(parts) == 3
        assert all(part.isdigit() for part in parts)


class TestMetadata:
    def test_declares_pytest_plugin_entry_point(self):
        with PYPROJECT.open("rb") as handle:
            config = tomllib.load(handle)
        entry = config["project"]["entry-points"]["pytest11"]["pyvbaharness"]
        assert entry == "pyvbaharness.pytest_plugin"

    def test_console_script_target_is_importable(self):
        from pyvbaharness.__main__ import main

        target = project_table()["scripts"]["pyvbaharness"]
        module, _, attribute = target.partition(":")
        assert module == "pyvbaharness.__main__"
        assert attribute == main.__name__

    def test_readme_and_license_are_present(self):
        project = project_table()
        assert (ROOT / project["readme"]).is_file()
        for pattern in project["license-files"]:
            assert (ROOT / pattern).is_file()

    def test_windows_only_dependency_is_markered(self):
        """pywin32 must stay behind a platform marker so the sdist can be
        resolved on non-Windows machines (pip download, mirrors, CI)."""
        deps = project_table()["dependencies"]
        pywin32 = [d for d in deps if d.startswith("pywin32")]
        assert pywin32 and 'sys_platform == "win32"' in pywin32[0].replace(
            "'", '"')

    def test_no_license_classifier_alongside_spdx(self):
        """PEP 639: an SPDX license field and a License classifier together
        are rejected by setuptools."""
        project = project_table()
        assert isinstance(project["license"], str)
        assert not any(c.startswith("License ::")
                       for c in project["classifiers"])
