# SIMS Doctor Site Diagnosis v0.4.0-RC6 Single-Code

RC5-HF-SITE-WIDE-PRECISION-RESULT の現行 Apps Script ロジックを、利用者が更新しやすい `Code.gs` 1本へ集約した構成整理版です。

## Apps Script投入対象

- `Code.gs`
- `appsscript.json`（権限・ランタイム変更時のみ更新）

## 方針

診断ロジック・メニュー・Evidence処理・Site-wide診断・Treatment Plan・Merge紹介状生成など、従来31個の `.gs` に分かれていた現行処理を `Code.gs` に統合しています。

RC5系からの初回移行時だけ、旧分割 `.gs` をすべて削除してから `Code.gs` を投入してください。以後の通常Hotfixは原則 `Code.gs` の置換だけで適用できます。

詳細は `RC6-SINGLE-CODE-APPLY.md` を参照してください。
