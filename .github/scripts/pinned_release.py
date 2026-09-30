"""Pinned GitHub release downloads: fetch them, and propose their updates.

    python pinned_release.py fetch <name> <dest-dir>
    python pinned_release.py propose [--cooldown-days N]

The pins live in .github/security/pins.json: a repository, a release tag,
the asset, and the SHA-256 the asset must have. Dependabot cannot follow a
release download, so ``propose`` does its job for these: it finds the newest
release at least ``--cooldown-days`` old, takes the asset's SHA-256 from the
digest GitHub recorded when the asset was uploaded, downloads the asset to
check it against that digest, and rewrites the pin. The weekly updater
workflow opens a pull request with the result, and the Security workflow
scans that pull request before it can be merged.

The digest always comes from GitHub's record, never from hashing the file
just downloaded: hashing what arrived proves only that it arrived.

Standard library only. GITHUB_TOKEN, when set, is used for API calls.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import sys
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path

PINS = Path(__file__).resolve().parents[1] / "security" / "pins.json"

# Every URL opened here is built from pins.json or read from GitHub's API.
# urllib also opens file:// and other schemes, so anything that is not
# HTTPS to GitHub is refused before it is opened.
ALLOWED_HOSTS = {"api.github.com", "github.com"}


def _request(url: str, accept: str = "application/vnd.github+json"):
    parts = urllib.parse.urlsplit(url)
    if parts.scheme != "https" or parts.hostname not in ALLOWED_HOSTS:
        raise ValueError(f"Refusing to open {url!r}: only HTTPS to "
                         f"{', '.join(sorted(ALLOWED_HOSTS))} is expected.")
    request = urllib.request.Request(url, headers={
        "Accept": accept, "User-Agent": "pyvbaharness-pin-updater"})
    token = os.environ.get("GITHUB_TOKEN")
    if token and parts.hostname == "api.github.com":
        request.add_header("Authorization", f"Bearer {token}")
    # The scheme and host were checked above; the rule matches any
    # non-literal URL, and this one cannot be a literal.
    return urllib.request.urlopen(request, timeout=120)  # nosemgrep: python.lang.security.audit.dynamic-urllib-use-detected.dynamic-urllib-use-detected


def _api(path: str):
    with _request(f"https://api.github.com/{path}") as response:
        return json.load(response)


def _download(url: str, target: Path) -> str:
    digest = hashlib.sha256()
    with _request(url, accept="application/octet-stream") as response, \
            open(target, "wb") as out:
        while chunk := response.read(1 << 20):
            digest.update(chunk)
            out.write(chunk)
    return digest.hexdigest()


def _asset_url(pin: dict) -> str:
    return (f"https://github.com/{pin['repo']}/releases/download/"
            f"{pin['tag']}/{pin['asset']}")


def fetch(name: str, dest: str) -> int:
    pin = json.loads(PINS.read_text(encoding="utf-8"))[name]
    target = Path(dest) / pin["asset"]
    target.parent.mkdir(parents=True, exist_ok=True)
    actual = _download(_asset_url(pin), target)
    if actual != pin["sha256"]:
        target.unlink(missing_ok=True)
        print(f"{name}: SHA-256 mismatch for {pin['asset']} at {pin['tag']}: "
              f"pinned {pin['sha256']}, downloaded {actual}. Refusing it.",
              file=sys.stderr)
        return 1
    print(f"{name}: {pin['repo']} {pin['tag']} {pin['asset']} verified "
          f"(sha256 {actual})")
    print(target)
    return 0


def _newest_eligible(repo: str, cooldown: dt.timedelta) -> dict | None:
    cutoff = dt.datetime.now(dt.timezone.utc) - cooldown
    for release in _api(f"repos/{repo}/releases?per_page=30"):
        if release.get("draft") or release.get("prerelease"):
            continue
        published = dt.datetime.fromisoformat(
            release["published_at"].replace("Z", "+00:00"))
        if published <= cutoff:
            return release
    return None


def propose(cooldown_days: int) -> int:
    pins = json.loads(PINS.read_text(encoding="utf-8"))
    cooldown = dt.timedelta(days=cooldown_days)
    changes: list[str] = []
    for name, pin in pins.items():
        release = _newest_eligible(pin["repo"], cooldown)
        if release is None:
            print(f"{name}: no release older than {cooldown_days} days")
            continue
        tag = release["tag_name"]
        if tag == pin["tag"]:
            print(f"{name}: {tag} is current")
            continue
        # A pin can be newer than anything past the cooldown, when a release
        # was adopted early on purpose. Never propose going backwards.
        if release["published_at"][:10] <= pin["published"]:
            print(f"{name}: pinned {pin['tag']} ({pin['published']}) is newer "
                  f"than the newest eligible release, {tag}; keeping it")
            continue
        asset_name = pin["asset_template"].format(tag=tag)
        asset = next((a for a in release.get("assets", [])
                      if a["name"] == asset_name), None)
        if asset is None:
            print(f"{name}: {tag} has no asset named {asset_name}",
                  file=sys.stderr)
            return 1
        recorded = asset.get("digest") or ""
        if not recorded.startswith("sha256:"):
            print(f"{name}: GitHub recorded no SHA-256 for {asset_name} at "
                  f"{tag}, so it cannot be verified. Not proposing it.",
                  file=sys.stderr)
            return 1
        expected = recorded.removeprefix("sha256:")
        with tempfile.TemporaryDirectory() as scratch:
            actual = _download(asset["browser_download_url"],
                               Path(scratch) / asset_name)
        if actual != expected:
            print(f"{name}: {asset_name} at {tag} does not match the digest "
                  f"GitHub recorded ({expected}); downloaded {actual}. "
                  "Not proposing it.", file=sys.stderr)
            return 1
        changes.append(f"{name}: {pin['tag']} -> {tag} "
                       f"({release['html_url']}), sha256 {expected}")
        pin.update(tag=tag, asset=asset_name, sha256=expected,
                   published=release["published_at"][:10])
    if not changes:
        print("Nothing to propose.")
        _output("changed", "false")
        return 0
    PINS.write_text(json.dumps(pins, indent=2) + "\n", encoding="utf-8")
    for line in changes:
        print(line)
    _output("changed", "true")
    _output("summary", "\n".join(changes))
    return 0


def _output(key: str, value: str) -> None:
    path = os.environ.get("GITHUB_OUTPUT")
    if not path:
        return
    with open(path, "a", encoding="utf-8") as out:
        out.write(f"{key}<<__END__\n{value}\n__END__\n")


def main(argv: list[str]) -> int:
    if len(argv) == 3 and argv[0] == "fetch":
        return fetch(argv[1], argv[2])
    if argv and argv[0] == "propose":
        days = 7
        if len(argv) == 3 and argv[1] == "--cooldown-days":
            days = int(argv[2])
        return propose(days)
    print(__doc__, file=sys.stderr)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
