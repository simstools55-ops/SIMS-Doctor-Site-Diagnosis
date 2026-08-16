# Release Notes — SIMS Doctor Site Diagnosis v0.8.0

## New workflow

The product now starts with whole-site analysis and then recommends one of two large flows:

- Existing articles: protect and repair current search traffic.
- Creator: pursue new article opportunities and content gaps.

The recommendation compares urgent A1/A2 and declining existing articles with detected
new-article opportunities and content gaps.

## Dialog-driven operation

Normal operation starts from:

`1. Site Diagnosisを進める`

The dialog shows the current priority route, current process step, required input/output,
the next operation, a route-switch button, and `中断して閉じる`.

## Resume

Closing the dialog does not reset the workflow. Route and current state are stored in
Document Properties. Existing package/result states remain the authoritative checkpoints,
so reopening the command resumes from the current process state.

## Compatibility

No Doctor result contract or SBM handoff contract is changed.
