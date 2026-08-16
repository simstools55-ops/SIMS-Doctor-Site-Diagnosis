# SIMS Doctor Site Diagnosis v0.4.0-RC6 Single-Code 適用手順

## 初回移行（RC5系 → RC6 Single-Code）

今回は分割 `.gs` を `Code.gs` 1本へ統合する構成変更です。

1. Apps Script プロジェクト内の既存 `.gs` ファイルをすべて削除します。
2. この配布物の `Code.gs` を追加します。
3. `appsscript.json` は既存内容と同一のため、通常は変更不要です。
4. 保存後、スプレッドシートを再読み込みし、メニューが表示されることを確認します。

旧 `.gs` と新 `Code.gs` を併存させないでください。

## 次回以降

通常更新は原則として以下です。

- 置換: `Code.gs`
- 変更なし: `appsscript.json`

HTMLファイルは現在ありません。
