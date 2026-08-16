# RC5 Final Display Fix

Freeze前の最後の表示修正。

- 「Doctor精密診断の優先候補」と「今回Doctorへ送る記事」の差を表示。
- 今回送らない優先候補数を表示。
- 既存データから確認できる場合は保護条件・高リスクを理由として表示。
- それ以外はTreatment Batchの選定上限・優先順位による見送りとして表示。

診断ロジック、Treatment Batch選定、Case Package、Hotfix 5分割再開、
Case Identity、request_idは変更しない。
