# SIMS Doctor Site Diagnosis v0.7.4 Patch

Base: v0.7.3

## Purpose

Improve completion visibility after Site Diagnosis results are registered in SBM.

## Changed

- After "SBMへの登録完了" is recorded, the Site Treatment Plan is refreshed immediately.
- Rows already handed to SBM are shown in gray.
- Their status is changed to "SBM引き渡し済み".
- NO_ACTION rows show "SBM引き渡し済み（処置不要）".
- NEEDS_EVIDENCE rows are intentionally left active and are not grayed out.
- The completed styling is re-applied whenever the treatment plan is regenerated, as long as the saved handoff fingerprint still matches the current payload.

## Unchanged

- Creator / Writer / Merge / Monitor routing
- Creator new-article identity semantics introduced in v0.7.3
- Evidence Package generation
- ZIP output folder selection
- Doctor result import
