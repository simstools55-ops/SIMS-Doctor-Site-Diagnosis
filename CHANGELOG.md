# v0.5.5

- Fixes duplicate treatment-plan rows after the final site-wide precision batch.
- Replaces unresolved NEEDS_EVIDENCE source cases by exact article URL when Doctor precision results omit the original case IDs.
- Uses exact diagnosis-theme matching only as a fallback when URL matching is unavailable.
- Restricts fallback removal to unresolved NEEDS_EVIDENCE cases to protect already finalized treatment cases.
- Preserves all v0.5.4 precision import and package-state behavior.

# v0.5.4

- Prioritizes a valid pasted Doctor site-wide/precision result over generating another package.
- Adds `WAITING_DOCTOR_RESULT` state for site-wide precision packages to prevent duplicate ZIP generation.
- Supports direct `clusters[].route_to` and `clusters[].articles` in `SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1`.
- Merges finalized precision clusters back into the stored site-wide diagnosis while preserving unresolved `NEEDS_EVIDENCE` cases.
- Uses absorbed source-case overlap as a safe fallback when replacing previously consolidated precision clusters.
- Renames the Doctor result import heading to clarify that the sheet is shared by site-wide and precision results.
- Migrates the import-sheet wording without clearing an already pasted Doctor result.
- Displays the generated precision ZIP filename in the completion dialog.
- Preserves v0.5.3 diagnosis/scoring logic, maximum five precision clusters per package, individual Doctor-to-SBM handoff, and Home version display.

# v0.5.3

- Unified Home/additional-Evidence counts and precision-package selection around the same `route_to = NEEDS_EVIDENCE` predicate.
- Removed the legacy `doctor_decision = ADDITIONAL_EVIDENCE_REQUIRED` hard filter from precision-cluster selection.
- Preserved the existing maximum of five precision clusters per package.
- Refreshes Home when no current precision cases exist so the UI does not retain a contradictory pending count.
- Preserved v0.5.2 individual Doctor-to-SBM handoff state, handed-off article exclusion, Home version display, and diagnosis/scoring behavior.

# v0.5.2

- Added the current product version to Diagnosis Home using `SDSD_VERSION`.
- Corrected the individual Doctor result handoff: individual case results return to SIMS-Blog-Manager, not Site Diagnosis.
- Added explicit confirmation that Doctor results were registered in SBM before marking the individual batch handed off.
- Records handed-off article URLs for the active diagnosis session and excludes them from subsequent individual precision selection.
- Keeps handed-off case rows for traceability while excluding them from pending-work counts.
- Clears individual Doctor package/wait/completed state when the diagnosis session ends.
- Preserved diagnosis/scoring logic, treatment rules, `maxPerRun = 3`, caching, and staged execution behavior.

# v0.5.1

- Fixed guided routing after individual Doctor Package export.
- Added `WAITING_DOCTOR_RESULT` state so `▶ 次に進む` does not regenerate the same Doctor Package while awaiting results.

# v0.5.0

- Promoted Site Diagnosis from the RC9.3.x line to the three-number product version scheme.
- Simplified normal operation around `Home`, `▶ 次に進む（Diagnosisに任せる）`, and the `確認する` menu.
- Moved direct workflow commands to `その他・管理 → 手動・保守操作`; retained them for recovery and maintenance.
- Removed duplicate user-facing menu entries that called the same individual precision diagnosis function.
- Updated Home, session, initialization, and completion guidance to the current guided-action menu model.
- Batched Case Enrichment sheet writes instead of repeated per-cell writes.
- Read Query Evidence once per Case Enrichment invocation and reuse both the map and source count.
- Batched Final Guard sheet/status writes.
- Preserved diagnosis/scoring algorithms, treatment rules, Article Identity semantics, caching, `maxPerRun = 3`, and staged resume behavior.
- Release ZIP is reduced to the current product files and excludes repository metadata, old Hotfix/Sprint/RC application notes, and development artifacts.

# v0.4.0-RC9.3.3

- Fixed stale title propagation in individual Doctor packages.
- Live fetched page title is now authoritative; Article Master title is fallback only.
- `article_url` is now the authoritative canonical identity used by Diagnosis.
- HTML `<link rel="canonical">` is stored separately as `observed_html_canonical_url`.
- Canonical mismatches no longer overwrite package identity; they are recorded as data-quality flags.
- Article Master/live-title mismatches are also recorded as data-quality flags.
- Applied the same canonical semantics to site-wide precision evidence.

# v0.4.0-RC9.3.2

- Removed manual Article Master CSV import as a prerequisite for individual Doctor precision packages.
- Article identity resolution is now automatic:
  1. Use real SBM ArticleID when available.
  2. Otherwise generate stable `REF-{URL hash}` internal ID.
- Canonical article URL is always retained as the authoritative fallback identity.
- Page title is fetched from the live article when Article Master title is unavailable.
- Main query falls back to the top query in Collector Evidence.
- Doctor package explicitly records whether ArticleID is real or URL-surrogate.
- SBM handoff can resolve surrogate cases by canonical URL.

# v0.4.0-RC9.3.1

- Fixed guided-route precedence: eligible A1/A2 individual precision candidates now take priority over stale/parallel site-wide additional-Evidence state.
- Home guidance now matches Site Diagnosis Detail for individual precision candidates.
- Added title fallback for Site Diagnosis Detail: Article Master -> live page title -> readable URL slug.
- Users still do not select rows manually.

# v0.4.0-RC9.3

- Added top-level `▶ 次に進む（Diagnosisに任せる）`.
- Diagnosis now automatically chooses the next workflow when user judgment is unnecessary.
- Individual precision diagnosis automatically creates the Treatment Batch and selects eligible A1/A2 articles.
- The same guided action continues through Article Master check, case enrichment, and Doctor Package generation.
- If Apps Script execution is split, pressing the same guided action resumes from the saved state.
- User action is requested only for missing prerequisites or genuine review cases.
- Site Diagnosis Detail now tells users to use the guided next action instead of selecting rows manually.

# v0.4.0-RC9.2

- Renamed `サイト診断サマリー` to `サイト診断詳細` with automatic sheet migration.
- Rebuilt the detail view with product-style colors and sections.
- Shows the actual highest-impact article titles and URLs instead of anonymous wording such as "影響の大きい2記事".
- Adds per-article symptom, cause to investigate, exact next action, and evidence/risk basis.
- Separates roles: Home = site overview / next action; Detail = causes / concrete diagnostic and treatment steps.
- Updated menu wording to `サイト診断詳細を見る`.

# v0.4.0-RC9.1

- Rebuilt Diagnosis Home as a dynamic site-diagnosis dashboard.
- Shows overall site status, candidate counts, decline signals, cannibalization, new-article opportunities and content gaps.
- Uses stored Doctor site-wide diagnosis as the highest-level site comment when available.
- Diagnosis automatically decides whether site-wide Doctor diagnosis or individual precision diagnosis is the next route.
- Home explains the next action, why it is needed, and the exact menu path.
- Removed user-facing wording that asked the user to decide "必要に応じて横断診断".
- Refreshes Home after Doctor site-wide result registration.

# v0.4.0-RC9

- Added Diagnosis Home with two primary workflows and next-action guidance.
- Reorganized menus around user workflow instead of internal functions.
- Replaced Evidence URL/file-ID input with a Collector-style Google Drive folder/file picker.
- Added pre-import Evidence metadata preview (site name, URL, generated time, collection period).
- Persist site_name/site_url from Collector manifest into the diagnosis session.
- Added `script.container.ui` OAuth scope for the Drive picker dialog.
- Preserved existing diagnosis/scoring logic and RC8 one-site session lifecycle.
- Existing `SIMS_DOCTOR_SITE_WIDE_RESULT_V1` and precision result support retained.

# v0.4.0-RC6 Single-Code

- RC5-HF-SITE-WIDE-PRECISION-RESULT の現行31個の `.gs` を `Code.gs` 1本へ統合。
- 診断ロジックの機能変更は行わず、更新・配布時の差し替え対象削減を目的とした構成整理。
- `appsscript.json` は変更なし。
- RC5系からの初回移行時は旧分割 `.gs` を削除してから適用する。

# Changelog

## v0.4.0-RC5 - Product UX

- Simplified the normal user menu to seven primary actions.
- Moved developer/regression validation commands under Maintenance.
- Hide internal diagnostic sheets during normal operation.
- Added lightweight progress Toast messages for long-running product flows.
- Added product wrappers for Site Diagnosis, Treatment Batch + Final Guard,
  and Case Enrichment + Doctor Case Package export.
- Rewrote completion messages for non-technical users.
- Preserved RC4 Case Identity, request_id, diagnosis scoring, and treatment
  selection logic without changes.

# Changelog

## 0.4.0-RC4

- Site Diagnosis Case Package に独立した `request_id` を必須発行。
- `case_id` / `request_id` / `site_diagnosis_case_id` / `site_diagnosis_batch_id` / `site_id` / `article_id` / `article_url` をトップレベルにも固定出力。
- 既存 `case_identity` は後方互換のため維持。
- ZIP manifest に各案件の `request_id` を追加。
- Export時に `request_id` 欠落をCase Identity不備としてブロック。
- `SDSD_VERSION` と `VERSION` の不整合を解消し、0.4.0-RC4へ同期。

## 0.4.0-RC3

- Recent Treatment GuardがSBM現行「改善履歴」の `改善実施日` を認識できず、直近処置済み記事をPASSしていた不具合を修正。
- 状態系ヘッダーの後方互換を維持しつつ、`判定` も参照候補へ追加。
- `_SDSD_SBM_HISTORY` の最新履歴をURL単位で照合し、直近35日以内の改善を `WAIT` にする既存仕様を復旧。
- Final Guardも同じ共通Guard関数を使用するため、Treatment Batch生成後の再確認にも修正が反映される。

## 0.4.0-RC2

- Doctor Case PackageのCase Identity欠落を修正。
- `site_diagnosis_batch_id` をTreatment Batch単位で発行。
- `site_diagnosis_case_id` / `individual_case_id` / `site_id` をCase Enrichmentで確定。
- SBM Article MasterのSiteIDを優先し、無い場合のみURLからsite_idを補完。
- Identity欠損時はDoctor Case Package ZIP生成を停止するfail-closed guardを追加。
- manifest contractを1.1へ更新し、Batch IDと各Case Identityを格納。

## 0.4.0-RC1

- Sprint 3 End-to-End実運用試験を完了。
- Collector Evidence Package ZIPのGoogle Drive直接取込を正式採用。
- `page_summary.csv` / `page_weekly.csv` / `page_query_top.csv` を自動展開。
- Query Evidenceを各Doctor案件へ最大10件付与。
- Query Evidenceに clicks / impressions / CTR / position を保持。
- Query Evidence欠損時のfail-closed判定を追加。
- 50,000文字セル制限を回避し、記事本文はZIP生成時に安全に添付。
- Site Size Policyによる治療件数上限を採用。
- SBM日常改善・モニター中記事とのTreatment Ownership重複防止を実装。
- Final Guardによる直近処置の再確認を実装。
- tonbos55 428記事規模で18案件のCase Package生成を実証。
- 本番メニューを整理し、Query Evidence診断を保守メニューへ移動。
- 過去Sprint/Hotfix適用文書をRelease ZIPから除外。

## 0.3.5

- Collector Evidence ZIPの直接取込を追加。

## 0.3.3

- Doctor Case PackageへのStructured Query Evidence格納を追加。

## 0.3.1

- SBM Article Master連携、本文取得、Case Enrichment、Batch ZIP exportを追加。

## 0.3.0

- Site Size Policy、Treatment Batch、Final Guardを追加。

## 0.2.x

- Weekly Trend、Evidence Confidence、Treatment Risk、Priority Gateを追加。

## 0.1.x

- URL正規化、TVS、Treatment Ownership、Recent Treatment Guard、候補一覧を実装。

## v0.4.0-RC8 - Diagnosis Session Lifecycle

- Added one-site-per-diagnosis-session lifecycle control.
- Added `診断セッション > 現在の診断状況を確認`.
- Added `診断セッション > 現在の診断を終了`.
- Blocked new Evidence imports while the current session remains active.
- Added explicit warning/counts before clearing current diagnostic work data.
- Session end clears Diagnosis working copies, including site-specific SBM history and Article Master copies, while leaving the original SBM untouched.
- No diagnosis scoring or treatment-selection logic changed.
