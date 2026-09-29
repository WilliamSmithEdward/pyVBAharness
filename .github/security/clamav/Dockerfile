# ClamAV engine for the malware scan in .github/workflows/security.yml,
# pinned by digest; Dependabot proposes version and digest bumps.
#
# The _base variant ships without signature databases. Signatures are not
# pinned: they change many times a day, so freshclam fetches them fresh on
# every run and checks their signatures, and the run's report records which
# database versions were used.
FROM clamav/clamav:1.5.4_base@sha256:7769870154c74ce31b0047dd8771e81f7c4269278bc005782e9e419e4922c73d
