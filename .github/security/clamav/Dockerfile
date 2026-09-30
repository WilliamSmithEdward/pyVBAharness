# ClamAV engine for the malware scan in .github/workflows/malware-scan.yml,
# pinned by digest; Dependabot proposes version and digest bumps.
#
# The same image every repository uses. Its bundled signatures are not used:
# the scan mounts its own database directory over /var/lib/clamav, and
# freshclam brings it up to date on every run and checks its signatures. The
# run's report records which database versions were used.
FROM clamav/clamav:1.5.4@sha256:ebec5bc138401b36ae987caa1a3fa3c3b2a21ed3d51f0bfa5852825e663e67b0
