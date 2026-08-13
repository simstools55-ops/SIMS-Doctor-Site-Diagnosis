# SIMS Doctor Site Diagnosis v0.4.0-RC4 適用手順

## 目的

Site Diagnosisが生成するDoctor Case Packageで `request_id` が欠落する問題を修正します。

## Apps Scriptで実際に置換するファイル

### 置換
- `CaseIdentity.gs`
- `CasePackageBuilder.gs`
- `SiteDiagnosisConfig.gs`

### 新規追加
- なし（Apps Script実行ファイルとして）

### 変更なし
- `Code.gs`
- `CaseSelection.gs`
- `RecentTreatmentGuard.gs`
- その他の `.gs` ファイル

## GitHubリポジトリで更新する管理ファイル

- `VERSION`
- `CHANGELOG.md`
- `README-FIRST.md`
- `RELEASE_NOTES_v0.4.0-RC4.md`（新規）

## バージョン

- Product: `SIMS Doctor Site Diagnosis`
- Previous: `0.4.0-RC3`
- New: `0.4.0-RC4`

`VERSION` と `SDSD_VERSION` はどちらも `0.4.0-RC4` に同期済みです。

## RC4で追加されるIdentity

Case Enrichment後の `case.json` は、トップレベルに以下を固定出力します。

- `case_id`
- `request_id`
- `site_diagnosis_case_id`
- `site_diagnosis_batch_id`
- `site_id`
- `article_id`
- `article_url`

`case_identity` は後方互換のため維持し、同じ `request_id` も保持します。

`request_id` 形式:

`REQ-<site_diagnosis_batch_id>-<ArticleID>`

例:

`REQ-20260813-213324-452426-A000107`

## UAT

1. 3つの `.gs` を置換し保存。
2. スプレッドシートを再読み込み。
3. 新しいTreatment Batchを1件以上生成。
4. Case Enrichmentを実行。
5. Doctor Case Package ZIPを生成。
6. `cases/.../case.json` を確認し、上記7 Identityがトップレベルにあることを確認。
7. Doctorへ渡し、返却JSONの `request_id` が `null` にならず、入力値をそのまま継承することを確認。
8. SBMの「Site Diagnosisの処置を進める」で登録し、同一Caseとして照合できることを確認。

## 推奨Git commit

`fix(site-diagnosis): issue and preserve request_id in RC4 case packages`
