# RC5 Product UX Hotfix 3

## Regression fix
- Restored/distributed `ArticleFetcher.gs` explicitly.
- Fixes `ReferenceError: sdsdFetchArticleEvidence_ is not defined` during menu 7.
- Case Package contract, request_id, Case Identity, scoring and treatment selection logic are unchanged.

## Product UX
- Added visible `サイト診断サマリー`.
- Candidate list now shows article title, URL, Japanese priority/reason and action.
- Candidate display is ordered by product priority first, then score.
- Removed `Doctor対象外` wording from the normal user view.
- Selected Treatment list now shows article title, Japanese reason and `サイト全体での意味`.
- Removed redundant visible `Doctor診断待ち` status column.
- Internal columns remain present but hidden.
- Added a priority legend to the summary and a note to the candidate priority header.
- Fixed literal `\n` appearing in RC5 completion dialogs.
- Re-hides Case Enrichment technical columns after menu 7.

## Scope boundary
Cannibalization, merge candidates and new-query/new-article discovery are intentionally not included in RC5 Hotfix 3.
They belong to the next Site Structure Diagnosis revision.
