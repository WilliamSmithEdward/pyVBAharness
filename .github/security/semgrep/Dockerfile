# Semgrep for .github/workflows/security.yml, pinned by digest;
# update-images.yml proposes version and digest bumps. Run from its image rather than pip,
# because its dependency tree is large enough that a hash-locked install
# would be most of the maintenance.
#
# The rulesets it applies (p/python and the rest) are fetched from the
# Semgrep registry on each run and are not pinned; the report names them.
FROM semgrep/semgrep:1.185.0@sha256:67518c174185a2043fdbed16c70631b79209499b3671178e7194379b48295a0d
