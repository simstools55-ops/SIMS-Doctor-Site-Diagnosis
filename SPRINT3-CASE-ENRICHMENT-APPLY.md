# Sprint 3.1 Case Enrichment

REPLACE:
- Code.gs
- SheetStore.gs
- SiteDiagnosisConfig.gs
- appsscript.json

NEW:
- ArticleMaster.gs
- ArticleFetcher.gs
- CasePackageBuilder.gs

ALL OTHER .gs FILES:
- NO CHANGE

After applying:
1. Save and reload.
2. If Google asks for new permissions, approve Drive + external URL access.
3. Run `1. Initialize` once so `_SDSD_ARTICLE_MASTER` exists.
4. Import the SBM `記事管理` CSV into `_SDSD_ARTICLE_MASTER`.
5. Run `12. Enrich Selected Cases`.
6. Confirm Ready / 要確認 counts.
7. Only when all selected cases are READY, run `13. Export Doctor Case Package ZIP`.

Important:
- This does not re-run TVS or treatment selection.
- Article fetch failures are not silently accepted.
