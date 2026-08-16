# Site Diagnosis v0.4.0-RC9 Product Workflow UX

## Home
Two primary workflows:
1. サイト全体を診断する
2. 見つかった問題を処置する

Home shows current site, Evidence file, candidate/pending counts and the next recommended action.

## Evidence input
Normal users no longer paste a Drive URL/file ID.
Use the Drive folder/file picker, inspect package metadata, then choose `このEvidenceを読み込む`.

## Color semantics
- Blue: normal workflow
- Green: completed/active confirmed state
- Yellow: confirmation or evidence required
- Red: error
- Gray: currently unnecessary

## Compatibility
Diagnosis/scoring logic is unchanged.
RC8 one-site session lifecycle remains active.
