# Semgrep for .github/workflows/security.yml, pinned by digest;
# update-images.yml proposes version and digest bumps. Run from its image rather than pip,
# because its dependency tree is large enough that a hash-locked install
# would be most of the maintenance.
#
# The rulesets it applies (p/python and the rest) are fetched from the
# Semgrep registry on each run and are not pinned; the report names them.
FROM semgrep/semgrep:1.182.1@sha256:451b3d875c00dac7ed5db5293e584a07de11449a409c10b3c2d2652f429630b6
