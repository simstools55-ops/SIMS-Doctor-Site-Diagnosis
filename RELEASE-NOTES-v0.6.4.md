# SIMS Doctor Site Diagnosis v0.6.4

## 修正内容

Creator候補チェックの表示時に、既存シートへ残った旧判定をそのまま表示してしまう問題を修正しました。

v0.6.4では「Creator候補チェックを見る」を開くたび、保存済みDoctor結果と現在のGSC Evidenceを使ってCreatorゲートを再評価し、シートを再生成します。

これにより、v0.6.3で追加した以下の判定が既存案件にも確実に反映されます。

- 候補KWそのものを既存記事が取得済み: RED / Creator除外
- similarity >= 0.85 の強近似クエリを既存記事が取得: YELLOW / Doctor SERP確認
- 既存担当記事・強近似担当が確認されない: GREEN

## Apps Script差し替え

- Code.gs: 置換
- appsscript.json: 変更なし
