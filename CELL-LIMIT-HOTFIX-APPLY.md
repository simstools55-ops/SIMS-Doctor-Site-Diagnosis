# Cell Limit Hotfix

REPLACE:
- CasePackageBuilder.gs
- SiteDiagnosisConfig.gs

ALL OTHER FILES:
- NO CHANGE

Fix:
- Article HTML is no longer embedded in `Referral JSON`.
- The sheet stores only lightweight metadata.
- Article HTML is kept temporarily in Document Cache when small enough.
- Export safely refetches the article when cache is unavailable/too large.
- ZIP contains `case.json` + `article.html` for each case.

Test:
1. Replace the two files.
2. Save and reload.
3. Run `12. Enrich Selected Cases` again. This overwrites the oversized Referral JSON values with lightweight JSON.
4. Confirm Ready 18 / 要確認 0.
5. Run `13. Export Doctor Case Package ZIP`.
