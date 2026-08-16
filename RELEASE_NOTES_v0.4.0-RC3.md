# SIMS Doctor Site Diagnosis v0.4.0-RC3

## Recent Treatment Guard Hotfix

### 原因
SBMの現行「改善履歴」CSVでは改善日の列名が `改善実施日` だが、Site Diagnosis RC2は
`改善日` / `処置日` / `日付` のみを参照していた。

そのためA000042の履歴自体はURLで見つかっていても日付が空扱いとなり、
Recent Treatment Guardが `PASS` を返していた。

### 修正
- `改善実施日` を改善日の第一候補として認識。
- `判定` を状態系の互換候補として追加。
- 旧ヘッダーは後方互換として維持。

### 回帰確認
A000042（2026/08/13改善）はSite Analysis再実行時に
Recent Treatment Guard = WAIT、Priority = WAIT となることを実運用で確認する。
