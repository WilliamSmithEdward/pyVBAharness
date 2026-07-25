# Releasing

Wheels are built and published by GitHub Actions. Authentication uses PyPI
Trusted Publishing, so no API token is stored in this repository or in
GitHub secrets: PyPI verifies a short-lived OIDC token that GitHub mints for
the workflow run, and trust is scoped to one repository, one workflow file,
and one environment.

## One-time setup

These steps need your PyPI account, so they have to be done by you in a
browser. The workflows are already written to match them.

### 1. Register the trusted publisher on PyPI

The project does not exist on PyPI yet, so register a *pending* publisher.
At <https://pypi.org/manage/account/publishing/>, add:

| Field | Value |
| --- | --- |
| PyPI project name | `pyvbaharness` |
| Owner | `WilliamSmithEdward` |
| Repository name | `pyVBAharness` |
| Workflow name | `publish.yml` |
| Environment name | `pypi` |

The first successful publish converts the pending publisher into a real
project owned by your account.

### 2. Create the GitHub environment

In the repository, under Settings > Environments, create `pypi`. The name
must match the workflow exactly, because PyPI checks the environment claim
in the OIDC token.

Adding required reviewers to that environment is worth considering: it turns
every publish into an explicit approval, which is a useful brake on an
accidental release.

## Cutting a release

1. Update the version in **both** places, which must agree:
   - `pyproject.toml`, the `version` field
   - `src/pyvbaharness/__init__.py`, `__version__`

   `tests/unit/test_packaging.py` fails if they drift, and the publish
   workflow refuses to build when the release tag disagrees with them.

2. Run the full local validation, including the live suite. Hosted runners
   have no Excel installation, so CI cannot do this for you:

   ```powershell
   python -m pytest tests/unit -q
   python -m pytest tests/live -m live -o addopts="" -q
   ```

   Confirm the live run left nothing behind:

   ```powershell
   Get-Process EXCEL -ErrorAction SilentlyContinue
   Get-ChildItem "$env:LOCALAPPDATA\pyvbaharness\sessions"
   ```

3. Refresh the benchmark baselines if anything touched the run path, and
   update the numbers quoted in the README:

   ```powershell
   python benchmarks/run_benchmarks.py --out benchmarks/output/baseline-<version>.json
   python benchmarks/run_pool_benchmarks.py --out benchmarks/output/pool-baseline-<version>.json
   ```

4. Commit, then dry-run the build: Actions > Publish > Run workflow. A
   manual run stops after the build job without publishing anything, and
   leaves the distributions as a downloadable artifact.

   Download it and test the exact file that would be published, on a machine
   that has Excel:

   ```powershell
   python -m venv .rc
   .rc\Scripts\python -m pip install (Get-ChildItem dist\*.whl)
   .rc\Scripts\pyvbaharness doctor --live
   ```

   This is worth more than a staging index would be. The artifact is
   byte-identical to what a release would upload, and it can be exercised
   against real Excel, which no hosted index or runner can do.

5. Tag and publish. Creating the GitHub release triggers the real publish:

   ```powershell
   git tag v<version>
   git push origin v<version>
   gh release create v<version> --title "v<version>" --notes "..."
   ```

## What the workflows do

`ci.yml` runs on pushes to main and on pull requests. It runs the unit suite
against Python 3.10 through 3.13 on Windows, builds the distributions, and
then verifies the wheel: that it is `py3-none-any`, that it contains the
package modules, that it does not ship the test suite, and that it imports
and exposes the console entry point from a clean virtual environment.

`publish.yml` runs the unit suite, checks the tag against the project
version, builds, verifies the wheel installs and imports, and uploads it.
Only a published release goes on to the PyPI job; a manual run stops after
the build, which is the dry-run path. The build job runs on Windows so the
tests are meaningful; publishing runs on Linux because it only moves files.

## Limits worth knowing

CI cannot validate against real Excel. Hosted runners have no Office
installation, so the 58 live tests, which are the ones that actually prove
the harness works, only ever run on a developer machine. Treat a green CI
badge as "the pure logic is intact and the package builds", not as "the
harness works".

The wheel is `py3-none-any` because the package is pure Python. It installs
on any platform, but only imports on Windows: several modules bind
`user32`/`kernel32` through ctypes at import time. The `pywin32` dependency
carries a `sys_platform == "win32"` marker so the metadata stays resolvable
on other platforms.

## If a publish fails

`Trusted publishing exchange failure` means PyPI did not accept the OIDC
token. Check that the environment name in the workflow matches the one
registered on PyPI exactly, that the workflow filename matches, and that the
repository owner and name match.

`File already exists` means that version was published before. PyPI does not
allow re-uploading a version, even after deleting it. Bump the version and
release again.
