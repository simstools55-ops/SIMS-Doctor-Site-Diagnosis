# Sprint 3 Apply

Replace:
- Code.gs
- SheetStore.gs
- SiteDiagnosisConfig.gs

New:
- SiteSizePolicy.gs
- CaseSelection.gs
- FinalGuard.gs

All other .gs files: NO CHANGE.
appsscript.json: NO CHANGE.

Test:
1. Replace/add the six files.
2. Save and reload.
3. Run `8. Build Treatment Batch`.
4. For the current 428-article tonbos55 dataset:
   - site-size policy should report Standard 12-18, Max 20.
   - only A1/A2 + DOCTOR_OWNED + PASS cases are eligible.
   - selection should never be padded with B/REVIEW/etc.
5. Run `9. Run Final Guard`.
6. Open `Selected Treatment Cases`.
