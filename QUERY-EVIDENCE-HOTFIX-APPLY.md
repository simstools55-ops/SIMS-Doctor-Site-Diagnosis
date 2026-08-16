# Query Evidence Hotfix

REPLACE:
- QueryEvidence.gs
- CasePackageBuilder.gs
- SiteDiagnosisConfig.gs

ALL OTHER FILES:
- NO CHANGE

Test:
1. Replace the three files.
2. Save and reload.
3. Run `12. Enrich Selected Cases` again.
4. Expected current tonbos55 batch: Ready 18 / 要確認 0.
5. `Query Evidence Count` should be greater than 0 for all 18 cases.
6. Run `13. Export Doctor Case Package ZIP`.
7. Upload the new ZIP for final inspection.
