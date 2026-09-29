"""Gate and summarise the security scans run by .github/workflows/security.yml.

    python security_report.py gate <path> [<path> ...]
    python security_report.py report <results-dir> <output.md>

``gate`` exits 1 when any SARIF file under the given paths carries a result.
CodeQL's analyze step uploads its findings but succeeds regardless, so
without this a release would ship past an alert nobody looked at.

``report`` writes the Markdown report attached to each GitHub release. It
reads whatever the scan jobs left behind and records a job that produced
nothing as exactly that, rather than as a pass.

Standard library only, so it runs on any runner without an install step.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
from pathlib import Path


def _sarif_files(paths: list[str]) -> list[Path]:
    found: list[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.is_file() and path.suffix == ".sarif":
            found.append(path)
        elif path.is_dir():
            found.extend(sorted(path.rglob("*.sarif")))
    return found


def _sarif_findings(path: Path) -> tuple[str, list[dict]]:
    """Tool label and one dict per result in a SARIF file."""
    data = json.loads(path.read_text(encoding="utf-8"))
    findings: list[dict] = []
    label = ""
    for run in data.get("runs", []):
        driver = run.get("tool", {}).get("driver", {})
        version = driver.get("semanticVersion") or driver.get("version") or ""
        label = f"{driver.get('name', '?')} {version}".strip()
        for result in run.get("results", []):
            location = (result.get("locations") or [{}])[0]
            physical = location.get("physicalLocation", {})
            findings.append({
                "rule": result.get("ruleId", "?"),
                "level": result.get("level", "warning"),
                "file": physical.get("artifactLocation", {}).get("uri", "?"),
                "line": physical.get("region", {}).get("startLine", "?"),
                "message": (result.get("message", {}).get("text", "")
                            .splitlines() or [""])[0],
            })
    return label, findings


def gate(paths: list[str]) -> int:
    files = _sarif_files(paths)
    if not files:
        # A scan that wrote nothing did not pass; it did not run.
        print("No SARIF output found to gate on.", file=sys.stderr)
        return 1
    total = 0
    for path in files:
        label, findings = _sarif_findings(path)
        total += len(findings)
        for finding in findings:
            print(f"{finding['file']}:{finding['line']}: {finding['rule']}: "
                  f"{finding['message']}")
        print(f"{path.name}: {label}: {len(findings)} finding(s)")
    if total:
        print(f"{total} finding(s). Fix them, or suppress one in source with "
              "a comment saying why, before this can pass.", file=sys.stderr)
        return 1
    return 0


def _job(name: str) -> str:
    return os.environ.get(f"JOB_{name.upper()}", "unknown")


def _row(check: str, tool: str, status: str, count: object) -> str:
    return f"| {check} | {tool} | {status} | {count} |"


def report(results_dir: str, output: str) -> int:
    root = Path(results_dir)
    lines: list[str] = []
    detail: list[str] = []
    failed = False

    ref = os.environ.get("REF_NAME", "")
    sha = os.environ.get("COMMIT_SHA", "")
    run_url = os.environ.get("RUN_URL", "")
    when = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    lines.append(f"# Security report{': ' + ref if ref else ''}")
    lines.append("")
    lines.append(f"Commit `{sha}`, scanned {when}.")
    if run_url:
        lines.append(f"Produced by [this workflow run]({run_url}).")
    lines.append("")
    lines.append("| Check | Tool | Result | Findings |")
    lines.append("| --- | --- | --- | --- |")

    # CodeQL, one SARIF per language.
    codeql_files = sorted(root.glob("sec-codeql-*/*.sarif"))
    if not codeql_files:
        lines.append(_row("CodeQL", "-", f"no output (job {_job('codeql')})",
                          "-"))
        failed = True
    for path in codeql_files:
        label, findings = _sarif_findings(path)
        language = path.parent.name.removeprefix("sec-codeql-")
        status = "pass" if not findings else "FAIL"
        failed |= bool(findings)
        lines.append(_row(f"CodeQL, {language}", label, status,
                          len(findings)))
        for finding in findings:
            detail.append(f"- CodeQL `{finding['rule']}` at "
                          f"`{finding['file']}:{finding['line']}`: "
                          f"{finding['message']}")

    # Semgrep, from its JSON, which carries the version and scan warnings.
    semgrep = root / "sec-semgrep" / "semgrep.json"
    if semgrep.is_file():
        data = json.loads(semgrep.read_text(encoding="utf-8"))
        results = data.get("results", [])
        status = "pass" if not results else "FAIL"
        failed |= bool(results)
        lines.append(_row("Static analysis",
                          f"Semgrep {data.get('version', '?')}", status,
                          len(results)))
        for result in results:
            detail.append(f"- Semgrep `{result.get('check_id', '?')}` at "
                          f"`{result.get('path', '?')}:"
                          f"{result.get('start', {}).get('line', '?')}`")
    else:
        lines.append(_row("Static analysis", "Semgrep",
                          f"no output (job {_job('semgrep')})", "-"))
        failed = True

    # pip-audit, for the runtime dependency tree of the package itself.
    audit = root / "sec-pip-audit" / "pip-audit.json"
    if audit.is_file():
        data = json.loads(audit.read_text(encoding="utf-8"))
        dependencies = data.get("dependencies", [])
        vulnerable = [d for d in dependencies if d.get("vulns")]
        count = sum(len(d["vulns"]) for d in vulnerable)
        status = "pass" if not count else "FAIL"
        failed |= bool(count)
        audited = ", ".join(f"{d['name']} {d.get('version', '?')}"
                            for d in dependencies) or "none"
        lines.append(_row("Known vulnerabilities in dependencies",
                          "pip-audit", status, count))
        detail.append(f"- Dependencies audited: {audited}.")
        for dep in vulnerable:
            for vuln in dep["vulns"]:
                fixes = ", ".join(vuln.get("fix_versions", [])) or "none yet"
                detail.append(f"- {dep['name']} {dep.get('version', '?')}: "
                              f"{vuln.get('id', '?')}, fixed in {fixes}")
    else:
        lines.append(_row("Known vulnerabilities in dependencies",
                          "pip-audit",
                          f"no output (job {_job('dependencies')})", "-"))
        failed = True

    # ClamAV and YARA-X, summarised by malware_scan.py.
    malware = root / "sec-malware" / "malware.json"
    if malware.is_file():
        data = json.loads(malware.read_text(encoding="utf-8"))
        clam, yara = data["clamav"], data["yara"]
        by_engine = {"clamav": [], "yara-x": []}
        for finding in data["unexpected"]:
            by_engine[finding["engine"]].append(finding)
        accepted = {"clamav": 0, "yara-x": 0}
        for finding in data["accepted"]:
            accepted[finding["engine"]] += 1
        for key, check, tool, done in (
                ("clamav", "Malware signatures",
                 f"ClamAV {clam['engine']}", clam["completed"]),
                ("yara-x", "YARA rules",
                 f"YARA-X {yara['engine']}", yara["completed"])):
            found = by_engine[key]
            status = "pass" if done and not found else "FAIL"
            count = str(len(found))
            if accepted[key]:
                count += f" ({accepted[key]} accepted)"
            lines.append(_row(check, tool, status, count))
        failed |= bool(data["failed"])
        official = ", ".join(f"{name} {version}" for name, version
                             in sorted(clam["databases"].items()))
        detail.append(
            f"- ClamAV signatures, fetched fresh for this run: official "
            f"{official or 'versions not recorded'}, plus "
            f"{len(clam['unofficial_files'])} files from the free "
            f"Sanesecurity, InterServer and URLhaus feeds; "
            f"{clam['signatures']} signatures in all, over "
            f"{clam['files_scanned']} files.")
        detail.append(f"- YARA rules: YARA Forge {yara['rules_release']}, "
                      f"{yara['rules']} rules, pinned by SHA-256.")
        for finding in data["unexpected"]:
            detail.append(f"- {finding['engine']} `{finding['match']}` in "
                          f"`{finding['path']}`")
        for finding in data["accepted"]:
            detail.append(f"- Accepted: {finding['engine']} "
                          f"`{finding['match']}` in `{finding['path']}`, "
                          "documented in "
                          "`.github/security/accepted-findings.toml`.")
        for match in data["unused_acceptances"]:
            detail.append(f"- Accepted finding `{match}` no longer matches "
                          "anything and can be removed.")
        for problem in data["problems"]:
            detail.append(f"- {problem}")
    else:
        for check, tool in (("Malware signatures", "ClamAV"),
                            ("YARA rules", "YARA-X")):
            lines.append(_row(check, tool,
                              f"no output (job {_job('malware')})", "-"))
        failed = True

    lines.append("")
    lines.append("**Overall: " + ("FAIL" if failed else "pass") + ".** "
                 "Any finding, or any scan that produced no output, fails "
                 "the release before it reaches PyPI.")
    if detail:
        lines.append("")
        lines.append("## Detail")
        lines.append("")
        lines.extend(detail)
    lines.append("")
    lines.append("Rulesets: CodeQL `security-extended` for Python and GitHub "
                 "Actions; Semgrep `p/python`, `p/security-audit`, "
                 "`p/secrets` and `p/github-actions`; ClamAV and YARA-X over "
                 "the source tree, the built wheel and sdist, and their "
                 "contents. Secret scanning with push protection runs on the "
                 "repository continuously. To report a vulnerability, see "
                 "SECURITY.md.")
    lines.append("")

    Path(output).write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines))
    return 1 if failed else 0


def main(argv: list[str]) -> int:
    if len(argv) >= 2 and argv[0] == "gate":
        return gate(argv[1:])
    if len(argv) == 3 and argv[0] == "report":
        return report(argv[1], argv[2])
    print(__doc__, file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
