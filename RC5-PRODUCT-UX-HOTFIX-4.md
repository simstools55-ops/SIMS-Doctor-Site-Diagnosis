# RC5 Product UX Hotfix 4

## Root causes
1. Menu 4 / 6 only opened existing sheets, so Hotfix 3 display changes were not applied to already-created RC5 sheets.
2. The 12 Case Package candidates failed because Case Enrichment requires the current site's SBM Article Master mapping, including ArticleID. This is expected contract protection, not a scoring failure.

## Fixes
- Menu 4 refreshes the visible Candidate view from preserved technical columns before opening it.
- Menu 6 refreshes the visible Selected Treatment view from preserved technical columns before opening it.
- Menu 7 performs Article Master coverage preflight before article fetching.
- When Article Master coverage is incomplete, menu 7 stops safely and shows:
  - selected case count
  - URL match count
  - ArticleID-ready count
  - exact next action
- Article Master import help explicitly says to use the current blog's SBM Article Management CSV.
- Existing Treatment Batch can be reused after importing Article Master; menus 1-5 do not need to be rerun.

No diagnosis scoring, request_id, Case Identity, Treatment Batch selection, or Case Package contract was changed.
