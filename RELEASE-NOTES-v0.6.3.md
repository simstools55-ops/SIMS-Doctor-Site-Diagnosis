# SIMS Doctor Site Diagnosis v0.6.3

Creator候補の事前判定が過大にGREENとなる問題を修正するPATCHです。

## 修正内容

- 候補キーワードそのものを既存記事がGSCで取得している場合はREDとし、Creator候補から除外します。
- 類似度0.85以上の強近似クエリを既存記事が取得している場合はYELLOWとし、DoctorのSERP確認を必須にします。
- GREENは既存担当記事が確認されない場合に限定します。
- 判定理由に実際のGSC Evidenceを表示します。
- Creator候補チェックの通常表示はGREEN / YELLOWのみです。
- Doctor結果取込後のHome「新規記事機会」はGREEN + YELLOW件数を表示します。

## Apps Scriptで置換するファイル

- `Code.gs` のみ

`appsscript.json` と `SITE-WIDE-RESULT-CONTRACT-V1.md` は変更ありません。
