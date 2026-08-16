# Changelog

## v0.7.2
- Prioritizes finalized SBM handoff before remaining Creator SERP-review candidates.
- Adds SBM handoff payload fingerprinting so new actionable cases invalidate an older completion state.
- Migrates older sessions safely: a legacy `COMPLETE` state without a fingerprint requires one fresh handoff.
- After handoff completion, unresolved Creator candidates resume normally.
- Updates handoff UI wording to include Creator alongside Writer / Merge / monitoring.
