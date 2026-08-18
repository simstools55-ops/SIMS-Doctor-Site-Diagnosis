# SIMS Doctor Site Diagnosis v0.10.1

## v0.10.1 — Real-site scale / routing hotfix

Changed runtime file: `Code.gs`

- Doctor results longer than 50,000 characters are automatically split across cells and reconstructed for import.
- Site-wide batch IDs no longer carry over to the next blog.
- Doctor return destination is explicitly Diagnosis, not SBM.
- Site-wide package filenames identify From→To, blog name, purpose, and timestamp.
- The next-blog session-end menu wording is clearer.

Replace `Code.gs`, update `VERSION`, save, and reopen the spreadsheet. `appsscript.json` is unchanged.


Changed runtime file: Code.gs

- Preserve Doctor precision-result allowed_scope / blocked_scope from cluster or root workflow_handoff.
- Refuse to generate an unsafe Writer handoff when allowed_scope is still missing.
- Fill site_id / site_name / site_url from the imported article master when the stored Doctor result leaves them blank.
- Precision Doctor referral now explicitly requires Writer treatment scopes.

Replace Code.gs in the Apps Script project, save, and reopen the spreadsheet.

## v0.10.0

- Home「サイト全体の分析結果」コメント欄の表示領域を拡張し、見切れを改善。
- `site_id` / host 単位で Diagnosis Session History v1 を保存。
- 新しいCollector Packageを読み込んだ際、同一サイトの前回診断履歴をHomeへ表示する土台を追加。
- 診断終了時に候補記事のスナップショットを内部履歴へ保存。
- 詳細な前回差分（改善／継続悪化／新規悪化）の自動算出は次段階の機能として未実装。