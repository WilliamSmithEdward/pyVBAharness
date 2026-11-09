# Semgrep for .github/workflows/security.yml, pinned by digest;
# update-images.yml proposes version and digest bumps. Run from its image rather than pip,
# because its dependency tree is large enough that a hash-locked install
# would be most of the maintenance.
#
# The rulesets it applies (p/python and the rest) are fetched from the
# Semgrep registry on each run and are not pinned; the report names them.
FROM semgrep/semgrep:1.183.0@sha256:2e45517a85e13a9cca14a2a00df6bb48affd66b8ebfae1b0991cd57578362fb0
