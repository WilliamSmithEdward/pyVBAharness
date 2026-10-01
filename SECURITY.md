# Security policy

## Reporting a vulnerability

Report a vulnerability privately, not in a public issue or pull request:
[open a private report](https://github.com/WilliamSmithEdward/pyVBAharness/security/advisories/new).
Only the maintainer sees it. Include the `pyvbaharness` version, the
Office and Windows versions, what you did, what happened and what should
have happened, and the smallest file or steps that show it, with
credentials and private data removed. A reproduction script is the most
useful thing you can include; the harness is easy to drive from a few
lines of Python, so most problems can be shown that way.

A confirmed vulnerability is fixed in a release on PyPI, and the advisory
is published with it, crediting you unless you ask otherwise.

## Supported versions

Only the latest release on PyPI receives security fixes. Older releases are
not maintained separately; update when a fix ships.

## Scope

The harness does a few things on your machine that are worth knowing
about, and they mark out what is and is not a vulnerability.

**It terminates processes.** When a run exceeds its deadline, the harness
kills the Office process it started, identified by process ID plus process
creation time so that a reused ID cannot match. It creates its own Office
instance and never attaches to one you have open. Any way to make it kill,
attach to, or run code in a process it did not create is in scope.

**It deletes registry entries.** After killing Office it removes
`Resiliency\DocumentRecovery` entries under `HKEY_CURRENT_USER`, and only
entries naming a document that session opened. Any way to make it remove an
entry it did not create, or any other key, is in scope.

**It writes files** under `%LOCALAPPDATA%\pyvbaharness`: session
manifests, which record the process IDs it may kill, and screenshots taken
when a run times out. Any way to plant or alter a manifest so the harness
kills something else is in scope.

**It makes no network connections** and sends no telemetry.

### Using it safely

**It runs VBA.** That is its purpose. Running VBA you do not trust through
the harness is running code you do not trust, the same as opening a
macro-enabled document from a stranger. The harness contains a hang or a
dialog; it does not sandbox what the code does.

**It needs "Trust access to the VBA project object model".** While that
setting is on, any code running as your Windows user can modify VBA
projects, not only the harness. That is why it belongs on a development
machine and not a shared server. The harness does not turn it on for you.

## How the code is checked

Three workflows check every pull request and every push to `main`, and
their gates decide whether a change can merge: **CI passed**,
**Security passed** and **Malware scan passed**. A gate passes only when
every job before it did, and any unexpected finding fails it, whatever its
severity. Security and Malware scan also run daily, and again from the
Publish workflow for every release.

- **Code:** CodeQL with the `security-extended` queries, for Python and the
  GitHub Actions workflows, and Semgrep with the `p/python`,
  `p/security-audit`, `p/secrets` and `p/github-actions` rule sets. Results
  go to the repository's code scanning.
- **Workflows:** zizmor audits the GitHub Actions workflows; a finding fails
  Security.
- **Dependencies:** pip-audit, on Windows, over the package's runtime
  dependencies as a user installing it today would resolve them. The only
  one, pywin32, installs only on Windows. A dependency that cannot be
  audited fails the job as well.
- **Malware:** ClamAV, with signatures freshclam fetches and verifies on
  every run, and YARA-X, with the YARA Forge rules pinned to a release and
  its SHA-256, scan the checkout (without `.git`), the built wheel and sdist,
  and their unpacked contents. ClamAV adds the free Sanesecurity,
  InterServer and URLhaus feeds to the official databases, and the scan
  fails if an official database or the unofficial feeds are missing. YARA-X
  runs the full YARA Forge pack. The report records the signature and rule
  versions used.
- **Fuzzing:** Atheris fuzzes the two parsers that read text the harness
  does not control (`fuzz/fuzz_harness.py`): the procedure signature parser,
  on arbitrary VBA source, and the worker protocol decoders, on arbitrary
  lines. Each starts from its seeds in `tests/fuzz_corpus`, which the unit
  suite also replays. The Fuzz workflow runs on every change to those two
  modules, the fuzz target or its corpus, for a minute per target, and
  daily for five. It is not a gate: a finding becomes a regression test
  with its fix. Its first run found that a worker line nested deeply enough
  raised `RecursionError` rather than `ValueError`, which would have stopped
  the session reading the worker's output; the decoders now refuse it as a
  garbled line.
- **OpenSSF Scorecard** rates the repository's security practices on every
  change to `main` and weekly, and the README badge shows the result.
  Some of its checks assume more than one maintainer, such as a second
  person approving every change, so a single-maintainer project cannot
  score full marks on them.

## Accepted findings

A malware finding is fixed, or accepted with a written reason in
[.github/security/accepted-findings.toml](.github/security/accepted-findings.toml).
An entry matches the engine, the exact signature or rule name, and the
paths it names, and must say why the finding is false and when the entry
can go; an entry without a reason fails the scan. An entry that no longer
matches anything fails the report, naming the entry and its reason, until
it is removed.
CodeQL has no accepted list: any finding fails Security. A Semgrep finding
can be accepted only by a `# nosemgrep: <rule-id>` comment at the line,
with a comment saying why. zizmor keeps its exceptions in
`.github/zizmor.yml` or inline beside the line they excuse, each with its
reason.

The current entries: no malware finding is accepted and no `nosemgrep`
comment is in use. zizmor has `self-repository` turned off in
[.github/zizmor.yml](.github/zizmor.yml) until GitHub's documentation
confirms the `$/` self-repository syntax for reusable workflows called from
`publish.yml`.

## Pinning and updates

Everything the workflows run is pinned: actions to full commit SHAs,
runners to named OS releases, scanner images to digests, Python tools to
hash-locked lock files, and the YARA-X engine and YARA Forge rules to a
release and its SHA-256. ClamAV's signatures change too often to pin, so
freshclam fetches and verifies them on every run. The unofficial ClamAV
feeds and Semgrep's registry rule sets are also fetched fresh on every run,
and the engines that read them are pinned. Runner labels and
Python versions are moved by hand. The test suite's own dependencies
resolve from the ranges `pyproject.toml` declares, because the point of the
test matrix is to show the package works with what a user would install.

Dependabot proposes updates to the GitHub Actions, the scanner images, the
package's dependencies in `pyproject.toml` and the lock files in
`.github/requirements` once a version is a week old, and at once for a
security advisory. The Update YARA rules workflow proposes new YARA pins
each week. A minor or patch update, and the YARA pull request, merges
itself once CI, Security and Malware scan pass; a third-party major version
waits for review.

## Releases

Publishing a GitHub release starts the Publish workflow. On Windows it runs
the unit tests, checks that the release tag matches the version in
`pyproject.toml`, builds the wheel and sdist, checks their metadata, and
installs the wheel in a clean environment to check that it imports. It runs
Security and Malware scan on the release commit, and uploads to PyPI
through trusted publishing only when the build and both scans pass, so no
upload token exists to be stolen. Started by hand, it is a dry run that
publishes nothing.

The reports are attached to the GitHub release as `security-report.md` and
`malware-report.md` whether the scans passed or not, so a blocked release
says what blocked it. A release that reached PyPI also carries the signed
provenance bundle `pyvbaharness-<version>.sigstore.json`.

### Verifying a download

Every file on PyPI carries PyPI's own provenance, which names this
repository's `publish.yml` as the publisher; the file's page on PyPI shows it.
Releases published after 2026-09-30 also carry a GitHub build provenance
attestation, which you can check against any copy of the file, from PyPI or
from the GitHub release:

```
pip download pyvbaharness --no-deps -d check
gh attestation verify check/<file> --owner WilliamSmithEdward
```

The output names the commit and workflow run that built the file. The
signed bundle is also attached to the GitHub release as
`pyvbaharness-<version>.sigstore.json`, so the check works without asking
GitHub for it: add `--bundle pyvbaharness-<version>.sigstore.json`.

## Repository settings

<!-- repo-standards:begin security-settings. Copied from WilliamSmithEdward/repo-standards, templates/security/settings-block.md. Change it there; the weekly rescan fails a copy that differs. -->
- `main` accepts changes only through a pull request that passes
  **CI passed**, **Security passed** and **Malware scan passed**. The
  ruleset has no bypass, for the owner either, and refuses force-pushes and
  deleting the branch.
- A `v*` release tag cannot be moved or deleted once pushed, except by a
  repository admin.
- A workflow that uses an action not pinned to a full commit SHA fails to
  run. Workflow tokens are read-only unless a job is granted more for
  itself.
- Secret scanning with push protection, Dependabot alerts and security
  updates, and private vulnerability reporting are on.
<!-- repo-standards:end -->
