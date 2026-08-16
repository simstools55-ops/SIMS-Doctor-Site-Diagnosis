# Site-Wide Precision Result Hotfix

## Purpose
Accept the second-stage Doctor return contract `SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1` without breaking the existing `SIMS_DOCTOR_SITE_WIDE_RESULT_V1` flow.

## Apps Script replacement
- REPLACE: `SiteWideResultContract.gs`
- REPLACE: `VERSION` only in repository/distribution tracking; VERSION is not an Apps Script source file.
- NO CHANGE: all other `.gs` files.

## Behavior
- Existing first-stage site-wide result remains backward compatible.
- Precision `clusters[]` are normalized into treatment-plan rows.
- `cluster_result.sub_groups[]` is flattened so one cluster can route some articles to Writer and others to Monitor.
- Precision routes accepted: WRITER / MERGE / MONITOR / NO_ACTION / NEEDS_EVIDENCE.
- Creator is intentionally not accepted in precision result, matching the precision referral contract.
