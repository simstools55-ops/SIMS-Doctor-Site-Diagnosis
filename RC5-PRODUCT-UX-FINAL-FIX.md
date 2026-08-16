# RC5 Product UX Final Fix

## Root cause
The latest candidate renderer existed in code, but an already-created `診断候補` sheet could remain in its old layout until the view was explicitly rebuilt.

## Fix
- Rebuild `診断候補` from its preserved hidden technical columns on spreadsheet open.
- Rebuild `今回の診断対象` on spreadsheet open.
- Menu 4 always refreshes the candidate display before opening it.
- `サイト診断サマリーを見る` reconstructs the summary from existing candidate data; Site Diagnosis does not need to be rerun.
- Re-distribute `SheetStore.gs` with the final candidate renderer to remove runtime-version uncertainty.

## Protected behavior
No changes to diagnosis scoring, Treatment Batch selection, Case Package generation,
Hotfix 5 resumable processing, Case Identity, or request_id.
