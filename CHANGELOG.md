# CHANGELOG

## v0.8.4

- Prevented already-completed SBM handoff cases from reviving after the same Doctor result is re-imported.
- Added a per-case completed handoff ledger keyed by diagnosis batch/case/route/Creator keyword/article URLs.
- Newly finalized cases remain handoff-eligible while previously completed Writer / Creator / Merge / Monitor cases are skipped.
- The completed-case ledger is cleared only by the existing diagnosis-session reset flow.

## v0.8.3

- Fixed Site-wide Precision Result scope loss during Diagnosis normalization.
- Root `workflow_handoff.allowed_scope / blocked_scope` now acts as a safe default for Writer clusters.
- Cluster/group scope values take precedence over root defaults.
- Added Diagnosis-side pre-handoff validation for Writer `allowed_scope`.
- Added Article Master fallback for site identity in SBM handoff JSON.
