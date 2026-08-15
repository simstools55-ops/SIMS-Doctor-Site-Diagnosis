# SIMS Doctor Site Diagnosis v0.5.0

SIMS Doctor Site Diagnosis は、Collectorで収集したEvidenceを読み込み、サイト全体の診断、個別精密診断、横断診断、Doctorへの引き渡しまでを案内する診断ワークスペースです。

## 通常の使い方

通常運用では、細かな処理メニューを選ぶ必要はありません。

1. スプレッドシートを開きます。
2. メニュー `SIMS Doctor Site Diagnosis` を開きます。
3. `▶ 次に進む（Diagnosisに任せる）` を実行します。
4. Homeの「次に行うこと」「理由」「操作」を確認します。
5. Apps Scriptの実行時間などで処理が分割された場合も、同じ `▶ 次に進む（Diagnosisに任せる）` を再度実行してください。

Diagnosisが現在の状態を確認し、Evidence Packageの読込み、サイト診断、個別精密診断、横断診断、Doctor結果登録など、次に必要な処理を自動判定します。

## 結果を確認するとき

`確認する` メニューから次の画面を開けます。

- `サイト診断詳細を見る`
- `診断候補を見る`
- `個別精密診断対象を見る`
- `サイト治療計画を見る`

処理を進めるときは `▶ 次に進む（Diagnosisに任せる）`、結果を見るときは `確認する` を使うのが基本です。

## Article Masterについて

Article Masterは通常運用では必須ではありません。

SBMのArticleIDが取得できる場合はそれを使用し、取得できない場合は記事URLを正本としてDiagnosis内部の安定IDを生成します。SBMへ戻す際もcanonical記事URLで照合できる設計です。

## 1サイトずつ診断します

Site Diagnosisは、1つのスプレッドシートで1サイトずつ診断する共有ワークスペースです。

別サイトへ移る前に、

`その他・管理 → 現在の診断を終了`

を実行してください。その後、`▶ 次に進む（Diagnosisに任せる）` から次のサイトの診断を開始できます。

長期的な改善履歴・治療後の経過観察は各サイトのSIMS-Blog-Managerで管理します。

## 手動・保守操作

通常は使用しません。トラブル対応や検証が必要な場合のみ、

`その他・管理 → 手動・保守操作`

から個別処理を実行できます。

## v0.5.0の主な変更

- 通常操作を `▶ 次に進む（Diagnosisに任せる）` 中心へ簡素化
- 確認用メニューを `確認する` に集約
- Case EnrichmentのSpreadsheet書込みを一括化
- Query Evidenceの重複読込みを削減
- Final GuardのSpreadsheet書込みを一括化
- Home・ダイアログの古いメニュー案内を現行操作へ統一
- 診断アルゴリズム、Treatment判定、`maxPerRun = 3`、分割・再開方式は変更なし

## 更新時に置き換えるファイル

既存のApps Script環境では、今回コード変更があるのは `Code.gs` です。

配布物の `VERSION`、`README-FIRST.md`、`CHANGELOG.md` もv0.5.0へ更新されています。`appsscript.json` と `SITE-WIDE-RESULT-CONTRACT-V1.md` は内容変更なしです。
