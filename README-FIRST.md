# SIMS Doctor Site Diagnosis v0.8.3

Changed runtime file: Code.gs

- Preserve Doctor precision-result allowed_scope / blocked_scope from cluster or root workflow_handoff.
- Refuse to generate an unsafe Writer handoff when allowed_scope is still missing.
- Fill site_id / site_name / site_url from the imported article master when the stored Doctor result leaves them blank.
- Precision Doctor referral now explicitly requires Writer treatment scopes.

Replace Code.gs in the Apps Script project, save, and reopen the spreadsheet.
