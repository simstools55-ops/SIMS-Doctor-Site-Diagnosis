# SIMS Doctor Site Diagnosis v0.7.1

PATCH for Creator long-tail workflow state propagation.

## Apps Script replacement
- Replace: `Code.gs`
- No other Apps Script files are required for this patch.

## Main fix
When a long-tail candidate is promoted to a Creator case, the parent Creator-validation case is treated as resolved and no longer reappears as a pending SERP-review candidate. Existing v0.7.0 sessions are self-repaired by detecting derived Creator cases already linked through `absorbed_source_case_ids`.
