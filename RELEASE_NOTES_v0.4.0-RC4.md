# SIMS Doctor Site Diagnosis v0.4.0-RC4

## Purpose

Site Diagnosis由来のDoctor Case Packageで `request_id` が欠落し、Doctorが `null` または `site_diagnosis_case_id` の代用を返す問題を修正します。

## Changes

- Case Enrichment時に案件固有の `request_id` を `REQ-<site_diagnosis_batch_id>-<ArticleID>` 形式で発行します。
- Case Packageのトップレベルへ以下のIdentityを固定出力します。
  - `case_id`
  - `request_id`
  - `site_diagnosis_case_id`
  - `site_diagnosis_batch_id`
  - `site_id`
  - `article_id`
  - `article_url`
- `case_identity` にも同じ `request_id` を保持し、既存Consumerとの後方互換を維持します。
- manifestにも `request_id` を記録します。
- ZIP生成時にrequest_idを含むIdentity完全性を検査します。

## Compatibility

既存のCase ID、Site Diagnosis Case ID、Batch IDの生成規則は変更しません。RC3までの既存案件のSBM登録互換性も変更しません。

## Apply

変更が必要なApps Scriptファイルは `CaseIdentity.gs`、`CasePackageBuilder.gs`、`SiteDiagnosisConfig.gs` です。
