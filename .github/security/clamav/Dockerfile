# ClamAV engine for the malware scan in .github/workflows/malware-scan.yml,
# pinned by digest; update-images.yml proposes version and digest bumps.
#
# The same image every repository uses. Its bundled signatures are not used:
# the scan mounts its own database directory over /var/lib/clamav, and
# freshclam brings it up to date on every run and checks its signatures. The
# run's report records which database versions were used.
FROM clamav/clamav:1.5.4@sha256:7054e3f34308c384ea43cd786ea17abf899bd44af40af9f1b099853d7a5c214e
