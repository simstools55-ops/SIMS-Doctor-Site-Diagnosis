# Hotfix: Site-wide Precision MERGE plan display

Apps Script replacement:
- Replace: SiteWideResultContract.gs
- New files: none
- No change: all other .gs files

Fix:
- Preserve `cluster_result.merge_plan`.
- Populate target article identities for MERGE cases.
- Add user-facing columns:
  - 統合先（残す記事）
  - 統合元（吸収する記事）
  - 統合方向
- Existing Precision Result validation and subgroup routing are unchanged.
