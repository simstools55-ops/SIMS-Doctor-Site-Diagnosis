# Sprint 2 Apply

Changed Apps Script files from Sprint1 Hotfix3:
- SiteAnalyzer.gs — REPLACE
- SheetStore.gs — REPLACE
- SiteDiagnosisConfig.gs — REPLACE

New Apps Script files:
- WeeklyTrendAnalyzer.gs — NEW
- EvidenceRisk.gs — NEW
- QueryEvidence.gs — NEW

All other .gs files: NO CHANGE.
appsscript.json: NO CHANGE.

Test:
1. Replace/add the six files.
2. Save and reload.
3. Run `4. Run Site Analysis`.
4. Confirm new columns:
   Weekly Trend / Evidence Confidence / Treatment Risk / External Factor.
5. Check REVIEW cases are either REVIEW or DOCTOR_REVIEW, not auto-promoted solely by TVS.
