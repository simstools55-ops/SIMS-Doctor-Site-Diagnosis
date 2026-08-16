# SIMS Doctor Site Diagnosis v0.6.5

## 修正内容

Creator候補のローカルゲートを追加調整しました。

- GSCに候補KWの完全一致・強近似クエリがなくても、Doctor一次診断の `target_articles` に既存記事URLがある場合はGREENにしません。
- その場合はYELLOW（SERP確認必須）として、既存記事との役割分担・検索意図の独立性をDoctorで確認します。
- GREENは、既存担当記事・強近似クエリ・同一テーマ既存URLのいずれも確認できない案件だけに限定します。

## Apps Scriptで置換するファイル

- `Code.gs`：置換
- `appsscript.json`：変更なし
