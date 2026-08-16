# RC5 HF4 - Merge Referral Plan-Only Source

## 変更理由
HF3/HF3aでは、Merge紹介状生成時に保存済みDoctor JSONを再読込・再検証していました。
しかし「サイト治療計画」には既にDoctorの最終判断と統合方向が登録済みです。

HF4では、保存済みDoctor JSONへの依存を廃止し、
「サイト治療計画」の選択行をMerge紹介状の正本として扱います。

## Apps Scriptへ反映するファイル

### 置換
- SiteWideMergeReferral.gs

### 新規追加
- なし

### 変更なし
- Code.gs
- SiteWideResultContract.gs
- SiteWidePrecisionPackage.gs
- SiteWideDoctorPackage.gs
- その他すべての .gs

## 安全仕様
- 選択行がDoctor判断=MERGEであることを必須確認
- 統合先（残す記事）を正本として保持
- 統合元（吸収する記事）を正本として保持
- 統合方向を再判定しない
- 紹介状生成のみ
- 記事本文変更・削除・301設定は行わない
