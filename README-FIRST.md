# SIMS Doctor Site Diagnosis v0.4.0-RC5

## 目的

SIMS Doctor Site Collector が生成したサイト全体の Evidence Package と、
SIMS-Blog-Manager（SBM）の改善履歴・記事情報を統合し、
Doctor が精密診断すべき記事をサイト規模に応じて選定して
Individual Doctor Case Package ZIP を生成します。

## 実証済みフロー

1. Site Collector の Evidence Package ZIP をGoogle Driveから直接読み込む
2. SBM改善履歴を反映する
3. サイト全体を分析する
4. 週次トレンド・Evidence Confidence・Treatment Riskを評価する
5. SBMの日常改善とDoctor治療の重複を防止する
6. サイト規模に応じてDoctor候補を絞り込む
7. Final Guardで直近処置を再確認する
8. SBM Article MasterでArticleID・タイトル・メインクエリを補完する
9. 記事本文を取得してCase Enrichmentする
10. 各案件へ上位10検索クエリを clicks / impressions / CTR / position 付きで付与する
11. Doctor Case Package ZIPを生成する


## v0.4.0-RC2 Case Identity Hotfix

- Treatment Batch作成時に `site_diagnosis_batch_id` を発行・保持します。
- 各Case Packageに `site_diagnosis_case_id` / `individual_case_id` / `site_id` を必須設定します。
- `site_id` はSBM Article MasterのSiteIDを最優先し、存在しない場合のみURLから決定論的に補完します。
- ZIP生成時にIdentity欠損をfail-closedで検出します。
- `manifest.json` にも3つのCase IdentityとBatch IDを保存します。
- DoctorはこれらのIDを生成せず、Case Packageの値をそのまま返却する前提です。

## v0.4.0-RC1 実運用検証

tonbos55（428記事規模）でEnd-to-End検証済み。

- Doctor選定案件: 18件
- Article Fetch Status: 全18件 VALID
- Case Package Status: 全18件 READY
- Query Evidence: 全18件に上位10クエリ
- Final Guard block: 0件
- 最終ZIP: manifest.json 1件 + case.json 18件 + article.html 18件 = 37ファイル
- 空本文・Query Evidence欠損・manifest不整合: 0件

## メニュー

1. 初期化
2. Evidence Package ZIPを読み込む
3. SBM改善履歴を取り込む
4. サイト分析を実行
5. 診断候補を開く
6. 週次トレンドを検証
7. 最終優先度を検証
8. 治療バッチを作成
9. Final Guardを実行
10. 選定案件を開く
11. Article Master取込案内
12. Case Enrichmentを実行
13. Doctor Case Package ZIPを生成

Query Evidence診断は保守メニューとして残しています。

## Evidence Package

Site Collector側でStep 5（ページ別上位クエリ）が完了したEvidence Packageを使用してください。
取込完了画面で `page_query_top` が0行の場合はサイト分析を続行せず、
Collector側のStep 5 / Evidence再生成を確認してください。

## 配布方針

このZIPはGitHub上書き用のクリーンRCです。
過去SprintのAPPLY文書・Hotfix適用メモは配布物から除外しています。
ファイル名はASCII英数字表記を使用しています。

## RC4 Identity Contract

Case Packageは `request_id` を含む7つの主要Identityをトップレベルへ固定出力します。Doctorはこれらを再生成せず、そのまま診断結果へ継承してください。



## RC5 Product UX

RC5は診断ロジックを変更せず、通常利用者向けUIを整理するリリース候補です。

- 通常メニューを7項目へ整理
- 開発・回帰検証機能を「保守・診断」へ隔離
- 内部処理用シートを通常は非表示
- サイト診断・Treatment Batch・Case Package生成に進捗Toastを追加
- 完了メッセージを利用者向け日本語へ整理
- Final GuardとCase Enrichmentを通常フロー内で自動実行

RC4のCase Identity / request_id / Treatment選定ロジックは変更していません。
