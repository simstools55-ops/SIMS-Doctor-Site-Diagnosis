# SIMS Doctor Site Diagnosis v0.4.0-RC5 Product UX

## Purpose
Turn the validated RC4 diagnosis engine into a cleaner end-user product without changing diagnosis logic.

## Product UX scope
- 7-item primary menu
- Data Preparation submenu
- Maintenance / Diagnostics submenu
- Internal-sheet visibility control
- Progress Toasts
- Japanese completion guidance
- Automatic Final Guard after Treatment Batch
- Automatic Case Enrichment before Doctor Case Package export

## Protected RC4 behavior
The following areas are intentionally unchanged:
- Case Identity contract
- request_id issuance/preservation
- Treatment Value Score
- Weekly Trend classification logic
- Evidence Confidence / Treatment Risk
- Site-size treatment capacity
- URL normalization and ownership rules

## Test status
Static validation only at build time.
Apps Script regression test is required before Freeze.
