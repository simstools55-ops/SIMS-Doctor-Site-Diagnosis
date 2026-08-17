# CHANGELOG

## v0.8.3

- Fixed Site-wide Precision Result scope loss during Diagnosis normalization.
- Root `workflow_handoff.allowed_scope / blocked_scope` now acts as a safe default for Writer clusters.
- Cluster/group scope values take precedence over root defaults.
- Added Diagnosis-side pre-handoff validation for Writer `allowed_scope`.
- Added Article Master fallback for site identity in SBM handoff JSON.
