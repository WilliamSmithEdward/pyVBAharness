# Security policy

## Reporting a vulnerability

Report it privately, through
[GitHub's private vulnerability reporting](https://github.com/WilliamSmithEdward/pyVBAharness/security/advisories/new).
Please do not open a public issue for anything you think is exploitable.

A useful report says which version you ran, what you did, what happened,
and what should have happened. A reproduction script is the most useful
thing you can include; the harness is easy to drive from a few lines of
Python, so most problems can be shown that way.

## Supported versions

Only the latest release. Fixes ship as a new release rather than being
backported, so upgrading is the fix.

## What counts

The harness does a few things on your machine that are worth knowing
about, and they mark out what is and is not a vulnerability.

**It runs VBA.** That is its purpose. Running VBA you do not trust through
the harness is running code you do not trust, the same as opening a
macro-enabled document from a stranger. The harness contains a hang or a
dialog; it does not sandbox what the code does.

**It needs "Trust access to the VBA project object model".** While that
setting is on, any code running as your Windows user can modify VBA
projects, not only the harness. That is why it belongs on a development
machine and not a shared server. The harness does not turn it on for you.

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

## What is checked automatically

Every push, every pull request, daily, and every release:

- **CodeQL**, with the `security-extended` queries, over the Python source
  and over the GitHub Actions workflows.
- **Semgrep**, with the `p/python`, `p/security-audit`, `p/secrets` and
  `p/github-actions` rulesets.
- **pip-audit**, for known vulnerabilities in the runtime dependencies.
- **ClamAV**, over the source tree, the built wheel and sdist, and their
  contents, with the official databases plus the free Sanesecurity,
  InterServer and URLhaus feeds. Signatures are brought up to date on
  every run, and their versions are recorded in the report.
- **YARA-X**, over the same files, with the full
  [YARA Forge](https://github.com/YARAHQ/yara-forge) rule package.

Any finding fails the run. On a release it stops publishing before PyPI,
and the report is attached to the GitHub release as
`security-report.md` whether the scans passed or not, so a blocked release
says what blocked it.

A malware finding known to be false is listed, with the evidence, in
[`.github/security/accepted-findings.toml`](https://github.com/WilliamSmithEdward/pyVBAharness/blob/main/.github/security/accepted-findings.toml).
Only the named rule in the named files is accepted. An entry without a
reason fails the scan, and one that stops matching is reported so it can be
removed. There is one at present: a YARA rule that fires on two
documentation files because of a YARA-X 1.20.0 engine fault, not because of
their content, described in full in that file.

## Pinning

Everything the scans and builds depend on is pinned, and every pin has
something that moves it:

| What | Pinned by | Moved by |
| --- | --- | --- |
| Workflow actions | commit SHA | Dependabot |
| ClamAV and Semgrep | image digest | Dependabot |
| Python tools for the scans | exact version and hash | Dependabot |
| `build` and `twine` for releases | exact version | Dependabot |
| YARA-X and its rules | release and SHA-256 | a weekly workflow |
| Runners and Python | OS release, exact version | by hand |

Dependabot and the weekly workflow both wait until a release is a week old
before proposing it. The weekly workflow takes each SHA-256 from the digest
GitHub recorded for the release asset, opens a pull request, and runs the
security scans on it before it can be merged.

Two inputs are deliberately left to change on their own: ClamAV's
signatures and Semgrep's registry rulesets, which change many times a day.
The engines that read them are pinned. The test suite's own dependencies
resolve from the ranges `pyproject.toml` declares, because the point of
that test matrix is to show the package works with what a user would
install.

Also on the repository: secret scanning with push protection, and
Dependabot alerts and security updates for vulnerable dependencies.

Releases reach PyPI through Trusted Publishing, so no upload token exists
to be stolen, and PyPI records an attestation tying each file to the
workflow run that built it.
