# Semgrep for .github/workflows/security.yml, pinned by digest;
# update-images.yml proposes version and digest bumps. Run from its image rather than pip,
# because its dependency tree is large enough that a hash-locked install
# would be most of the maintenance.
#
# The rulesets it applies (p/python and the rest) are fetched from the
# Semgrep registry on each run and are not pinned; the report names them.
FROM semgrep/semgrep:1.184.0@sha256:cc1b68266f267878b39739373ea9985b9a78e5d64497dd875c8de100b22695ed
