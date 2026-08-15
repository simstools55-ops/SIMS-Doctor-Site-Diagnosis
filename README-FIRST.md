# SIMS Doctor Site Diagnosis v0.5.5

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

## v0.5.5の主な変更

- 精密診断結果の統合時、元の `NEEDS_EVIDENCE` 案件を case ID だけでなく記事URLでも照合して置換
- URLが取れない場合のみ診断テーマの完全一致を安全なフォールバックとして使用
- 確定済みPrecision案件と元の追加Evidence案件がサイト治療計画へ二重表示される問題を修正
- 照合対象は未確定の `NEEDS_EVIDENCE` 案件に限定し、既存のWriter/Merge/NO_ACTION案件を誤削除しない
- v0.5.4の結果取込優先、Package重複生成防止、残案件保持、ZIPファイル名表示を維持

## v0.5.4の主な変更

- Doctor結果取込に有効な横断／精密診断JSONがある場合、Package生成より結果登録を最優先
- サイト横断精密診断Package生成後は結果待ち状態を保持し、同じPackageの重複生成を防止
- `SIMS_DOCTOR_SITE_WIDE_PRECISION_RESULT_V1` の直接 `clusters[].route_to` / `clusters[].articles` 形式に対応
- Precision Resultの確定クラスタを元の横断診断結果へ差し戻し、未診断の `NEEDS_EVIDENCE` 案件を残す
- Doctor結果取込シートを「横断診断／精密診断 共通」と明示し、貼り付け済み結果を消さない移行方式に変更
- 精密診断Package生成完了ダイアログにZIPファイル名を表示
- v0.5.3までのDoctor→SBM引き渡し、再選定防止、Homeバージョン表示、最大5クラスタ制限を維持
- 保存先フォルダー指定はCollector方式を確認してから別リリースで対応

## 更新時に置き換えるファイル

既存のApps Script環境では、今回コード変更があるのは `Code.gs` です。

配布物の `VERSION`、`README-FIRST.md`、`CHANGELOG.md` もv0.5.4へ更新されています。`appsscript.json` と `SITE-WIDE-RESULT-CONTRACT-V1.md` は内容変更なしです。
