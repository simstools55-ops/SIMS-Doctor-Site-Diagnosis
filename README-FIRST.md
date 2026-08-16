# SIMS Doctor Site Diagnosis v0.6.4

SIMS Doctor Site Diagnosis は、Collectorで収集したEvidenceを読み込み、サイト全体の診断、個別精密診断、横断診断、Doctorへの引き渡しまでを案内する診断ワークスペースです。

## 通常の使い方

通常運用では、細かな処理メニューを選ぶ必要はありません。

1. スプレッドシートを開きます。
2. メニュー `SIMS Doctor Site Diagnosis` を開きます。
3. `▶ 次に進む（Diagnosisに任せる）` を実行します。
4. 以後は、各ダイアログに表示される「何のための処理か」「完了したこと」「次にすること」と次操作ボタンに従います。
5. Apps Scriptの実行時間などで中断した場合や、現在地が分からなくなった場合だけ `▶ 次に進む（Diagnosisに任せる）` を再度実行してください。

Evidence Package読込後はそのままサイト診断へ進め、Doctor Package生成後はDoctor結果取込へ案内します。Doctor回答は通常、専用ダイアログへ全文またはJSONを貼り付けて取り込みます。

## 結果を確認するとき

`確認する` メニューから次の画面を開けます。

- `サイト診断詳細を見る`
- `診断候補を見る`
- `個別精密診断対象を見る`
- `サイト治療計画を見る`

処理を進めるときは `▶ 次に進む（Diagnosisに任せる）`、結果を見るときは `確認する` を使うのが基本です。


## Creatorルート（v0.6.4）

Site Diagnosisは、新記事候補を「作らないための審査」ではなく、重大なカニバリを避けながら有望なロングテールへ挑戦するための機会として扱います。

1. Doctor一次診断で `eventual_route = CREATOR` となった案件を、通常のPrecision診断案件から分離します。
2. Article MasterとGSCクエリを使い、既存記事との近さ、キーワードクラスター、差別化語、内部リンク候補を確認します。
3. 暫定判定は GREEN / YELLOW / RED。候補KWそのものを既存記事がGSCで取得していればRED、強い近似KWを取得していればYELLOW、既存担当記事が確認されない場合だけGREENとします。
4. 選択案件についてDoctorへSERP確認紹介状を作り、実SERP上の独立性を `CREATOR / WRITER / BLOCK / NEEDS_EVIDENCE` で確定します。
5. `CREATOR` 確定時は、メインKW、検索意図、既存記事との役割分担、狙わない意図、内部リンク候補、新記事作成理由、約30日のモニター条件を `creator_plan` としてSBMへ引き渡します。
6. 新記事公開後の実績判定はSBMで行い、データ不足ならMONITOR延長、カニバリが確認された場合はDoctor / Writer / Mergeへ戻します。

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

## v0.6.4の主な変更

- 候補KWと完全一致するGSCクエリを既存記事が取得している場合はREDとしてCreator候補から除外
- similarity 0.85以上の強近似クエリはYELLOWとしてDoctor SERP確認へ
- GREENを「既存担当記事が確認されない候補」に限定
- 判定理由へ既存URL・クエリ・表示回数・類似度の実Evidenceを表示
- Creator候補チェックはGREEN / YELLOWのみを通常表示
- Homeの「新規記事機会」を再判定後のGREEN + YELLOW件数で表示

## v0.6.2の主な変更

- Creator候補を通常の追加Evidence / Precision診断から分離
- Creator候補チェック（GREEN / YELLOW / RED）とGSCクエリクラスター・役割分担・内部リンク候補を追加
- 選択Creator案件のDoctor SERP確認紹介状と `SIMS_DOCTOR_CREATOR_SERP_RESULT_V1` 取込を追加
- CREATOR確定時に `creator_plan` をDiagnosis→SBM handoffへ追加
- Evidence読込、Doctor Package生成、Doctor結果取込をSBM型の「目的・完了・次操作」ダイアログへ改善
- 横断Diagnosis / Precision ZIP名にASCIIのサイト識別子を追加
- Precision Package生成済み時の重複生成を避け、Doctor回答取込へ誘導

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

配布物の `VERSION`、`README-FIRST.md`、`CHANGELOG.md` もv0.6.4へ更新されています。`appsscript.json` と `SITE-WIDE-RESULT-CONTRACT-V1.md` は内容変更なしです。

## v0.6.2 Creator SERP回答取込導線

Creator SERP確認後のDoctor回答は、`▶ 次に進む（Diagnosisに任せる）` または `確認する → Creator SERP Doctor回答を取り込む` から専用ダイアログへ貼り付けます。通常フローでSBMへ直接貼り付ける必要はありません。
