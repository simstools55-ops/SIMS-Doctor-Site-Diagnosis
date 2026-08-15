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
