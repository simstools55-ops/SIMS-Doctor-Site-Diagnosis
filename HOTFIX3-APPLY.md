# Hotfix3 Apply

Changed Apps Script files from Hotfix2:
- SiteAnalyzer.gs — REPLACE
- SiteDiagnosisConfig.gs — REPLACE

All other .gs files: NO CHANGE.
appsscript.json: NO CHANGE.

After applying:
1. Save and reload the spreadsheet.
2. Run `4. Run Site Analysis` again.
3. Confirm rows with Ownership = REVIEW have Priority Candidate = REVIEW.
4. Confirm WAIT, SBM, and DOCTOR_OWNED rows retain their previous behavior.
