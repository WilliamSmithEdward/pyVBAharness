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

2. Install the working tree as editable first, and check that it is what
   gets imported:

   ```powershell
   python -m pip install -e ".[dev,fuzz]"
   python -c "import pyvbaharness, pyvbaharness.__main__ as m; print(pyvbaharness.__file__); print(m.__file__)"
   ```

   This is not optional housekeeping. The package registers a `pytest11`
   entry point, so pytest imports `pyvbaharness` from site-packages at
   startup, before `tests/conftest.py` can put `src` on the path. With a
   non-editable install of a previous version present, the suite silently
   tests the installed package instead of the working tree, and a fix under
   test appears not to work. Both paths printed above must be inside `src`.

   Check a submodule as well as the package, because the two can disagree.
   Uninstalling leaves `site-packages\pyvbaharness\__pycache__` behind, and
   the surviving directory becomes a namespace package that wins over the
   editable finder for submodules while `import pyvbaharness` still resolves
   to `src`. If the second path is not in `src`, delete that directory:

   ```powershell
   Remove-Item "$(python -c 'import site;print(site.getsitepackages()[0])')\pyvbaharness" -Recurse -Force
   python -m pip install -e ".[dev,fuzz]"
   ```

3. Run the full local validation, including the live suite. Hosted runners
   have no Office installation, so CI cannot do this for you:

   ```powershell
   python -m pytest tests/unit -q
   python -m pytest tests/live -m live -o addopts="" -q
   ```

   Confirm the live run left nothing behind:

   ```powershell
   Get-Process EXCEL, WINWORD, POWERPNT, MSACCESS -ErrorAction SilentlyContinue
   Get-ChildItem "$env:LOCALAPPDATA\pyvbaharness\sessions"
   ```

   The live suite needs Excel, Word, PowerPoint and Access installed, and it
   opens PowerPoint on screen because PowerPoint cannot be hidden. Run it in
   the foreground; the whole suite takes 4 to 5 minutes.

4. Refresh the benchmark baselines if anything touched the run path, and
   update the numbers quoted in the README:

   ```powershell
   python benchmarks/run_benchmarks.py --out benchmarks/output/baseline-<version>.json
   python benchmarks/run_pool_benchmarks.py --out benchmarks/output/pool-baseline-<version>.json
   ```

5. Commit, then dry-run the build: Actions > Publish > Run workflow. A
   manual run is always a dry run: it publishes nothing, and leaves the
   distributions and both reports as the `release-preview` artifact.

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

   Check the Security run is green on the commit you are about to tag.
   Publishing will run it again as a gate, but finding out there costs a
   tag and a release to clean up.

6. Tag and publish. Creating the GitHub release triggers the real publish,
   behind the security gate, and attaches `security-report.md` to the
   release:

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

`security.yml` runs CodeQL (`security-extended`, over the Python and the
workflows), Semgrep (`p/python`, `p/security-audit`, `p/secrets`,
`p/github-actions`) and pip-audit. `malware-scan.yml` runs ClamAV and
YARA-X over the wheel and sdist, their contents, and the source tree. Each
fails on any finding and runs on pushes, pull requests and daily, and
`publish.yml` calls both as gates, so a finding stops a release before
anything is uploaded. Either way `security-report.md` and
`malware-report.md` are attached to the GitHub release, so a blocked
release says what blocked it. The dry-run path runs the scans too.

A finding is fixed, or recorded as false where the evidence supports it:

- Semgrep: `# nosemgrep: <rule-id>` in source, with a comment saying why.
- CodeQL: a dismissal on the alert in the Security tab, with a reason.
- ClamAV and YARA-X: an entry in `.github/security/accepted-findings.toml`
  naming the rule, the files, the evidence, and what would let the entry
  be removed. The scan refuses an entry without a reason, and fails on one
  that no longer matches anything.

There is no threshold to lower, deliberately.

Every pin the scans depend on has an updater; SECURITY.md has the table.
Dependabot proposes action, image and tool bumps weekly, each at least a
week old. `update-yara-rules.yml` does the same for YARA-X and its rules,
which are GitHub release downloads Dependabot cannot follow: it opens a
pull request pinning the new release and SHA-256, and dispatches Malware
scan and CI on that branch, because a pull request opened by a workflow
starts no workflows of its own. Merge it once they are green. Keep new
actions pinned the same way: Semgrep's `github-actions-mutable-action-tag` rule
fails the build on a `@v4`.

ClamAV brings its signatures up to date on every run, from a cache so that
only the day's changes are downloaded, and fails the scan if that update
fails, rather than scanning with whatever the cache held. If a run
fails at freshclam, it is usually ClamAV's mirrors rate-limiting the
runner; rerun it later.

Two repository rulesets (Settings > Rules > Rulesets) protect what a
release depends on.

"main" covers the default branch. A pull request needs "CI passed",
"Security passed" and "Malware scan passed" to merge, each pinned to the
GitHub Actions app so no other integration can satisfy them. Each gate job
depends on every other job in its workflow and passes only if all of them
succeeded, so renaming or adding a job, or moving a Python pin that names
a matrix job, never needs the ruleset changed. A new job does need adding
to its gate job's `needs`. Force pushes and
deleting the branch are refused.

"release tags" covers `v*`. Only a repository admin can create, move or
delete one, because pushing a tag and publishing a release from it is what
sends a version to PyPI.

Repository admins are the one bypass actor on both, so the maintainer can
still push to `main` directly (the checks then run after the push) and tag
releases as step 6 describes. To hold admins to the checks too, remove the
bypass entry from the "main" ruleset, or set it to pull requests only.

## Limits worth knowing

CI cannot validate against real Office. Hosted runners have no Office
installation, so the 129 live tests, which are the ones that actually prove
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
