# v0.10.3 MIXED Precision Routing — Apply

## Apps Script

Replace only:

- `Code.gs`

No change:

- `appsscript.json`

## What changed

Precision Doctor results may contain one cluster whose articles require different actions. Diagnosis now accepts:

- cluster `route_to: "MIXED"`
- mandatory `articles[].route_to`
- optional MERGE target fields on the MERGE article (`merge_target_url`, `merge_target_title`)

Diagnosis expands MIXED to concrete final treatment cases. MIXED is not retained as a final route.

## Regression test

Re-import the Doctor Precision Result that previously failed with:

`Precision Result の処置ルートが不正です: MIXED`

Expected:

- Teams cluster expands to MERGE + MONITOR.
- Excel cluster whose articles are all MONITOR resolves to MONITOR.
- The original `site_diagnosis_case_id` is preserved as the parent case ID.
