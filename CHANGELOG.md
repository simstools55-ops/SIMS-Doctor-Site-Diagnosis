# CHANGELOG

## v0.10.1

- Added chunked persistence for Doctor site-wide responses exceeding the 50,000-character Google Sheets cell limit.
- Prevented site-wide batch ID reuse across diagnosis sessions/sites.
- Added active-session fallback for site identity in site-wide Doctor packages.
- Corrected Doctor referral return destination to SIMS Doctor Site Diagnosis and made JSON output mandatory.
- Renamed generated site-wide Doctor packages to include From→To, readable site name, purpose, and timestamp.
- Clarified the menu action used to finish the current site and switch to the next blog.

## v0.10.0

- Added Diagnosis Session History v1 keyed by site_id/host.
- Added hidden `_SDSD_DIAGNOSIS_SESSIONS` and `_SDSD_DIAGNOSIS_CASE_HISTORY` stores.
- Added previous-session summary to Diagnosis Home.
- Expanded Home site-wide analysis comment display to prevent clipping.
- Preserved the existing Collector-is-stateless / Diagnosis-keeps-history responsibility split.

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
