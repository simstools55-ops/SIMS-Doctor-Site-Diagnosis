# Evidence ZIP Import

REPLACE:
- Code.gs
- SiteDiagnosisConfig.gs

NEW:
- EvidencePackageImporter.gs

ALL OTHER FILES:
- NO CHANGE

Test:
1. Replace/add the three files.
2. Save and reload.
3. Run `2. Import Evidence Package ZIP`.
4. Paste the Google Drive URL of the RC5 Collector Evidence ZIP.
5. Confirm the completion dialog shows non-zero rows for:
   - page_summary
   - page_weekly
   - page_query_top
6. Run `13. Diagnose Query Evidence Input`.
7. Expected for current tonbos55 Evidence:
   - Raw Rows > 0
   - Selected Cases Matched to Query Evidence = 18/18
8. Then run `12. Enrich Selected Cases` again.
9. Confirm Query Evidence Count > 0 for all 18.
10. Finally run `14. Export Doctor Case Package ZIP`.
