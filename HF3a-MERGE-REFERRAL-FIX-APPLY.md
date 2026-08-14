# RC5 HF3a - Merge Referral stored-result compatibility fix

## 症状
「15. 選択中のMerge紹介状を作成」で
「保存済みDoctor横断診断結果の形式が不正です。」
となる。

## 原因
HF3が保存済みDoctor結果を `diagnosis_cases[]` 形式だけで読む前提だった。
実運用で登録済みの精密診断結果は `clusters[]` / `sub_groups[]` を保持する形式もある。

## 修正
保存済み結果の両形式を受理する。
さらに、登録時にクラスタ詳細が正規化されて失われている場合でも、
「サイト治療計画」に既に確定保存された
- Doctor判断=MERGE
- 統合先（残す記事）
- 統合元（吸収する記事）
を正本としてMerge紹介状を生成する。

## Apps Scriptへ反映するファイル
### 置換
- SiteWideMergeReferral.gs

### 変更なし
- Code.gs
- SiteWideResultContract.gs
- その他すべての .gs
