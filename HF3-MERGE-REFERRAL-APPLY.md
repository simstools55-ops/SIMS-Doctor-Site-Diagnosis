# RC5 HF3 - Merge Referral

## Apps Scriptへ反映するファイル

### 新規追加
- SiteWideMergeReferral.gs

### 置換
- Code.gs

### 変更なし
- SiteWideResultContract.gs
- SiteWidePrecisionPackage.gs
- SiteWideDoctorPackage.gs
- CasePackageBuilder.gs
- ProductUX.gs
- その他すべての .gs

## 使い方
1. 「サイト治療計画」を開く
2. MERGE案件の行を1行選択
3. メニュー「15. 選択中のMerge紹介状を作成」
4. 「Merge紹介状」シートに生成されたMarkdownまたはJSONをSIMS Mergeへ渡す

## 安全仕様
- Doctorが確定した統合方向を再判定しない
- 記事本文の変更はしない
- 記事削除はしない
- 301リダイレクトは設定しない
- 紹介状生成だけを行う
