# Changelog

## v0.7.3
- Separates the true Creator new-article target from existing related/reference articles.
- Shows Creator treatment-plan targets as `新規記事（未発行） / キーワード: ...` instead of an existing related URL.
- Adds explicit `new_article_target`, `reference_articles`, and `article_identity_semantics` fields to the SBM handoff payload.
- Retains legacy reference-article identity data only for backward compatibility with current SBM validation.
- Aligns the repository `VERSION` file with the actual product version.

## v0.7.2
- Prioritizes finalized SBM handoff before remaining Creator SERP-review candidates.
- Adds SBM handoff payload fingerprinting so new actionable cases invalidate an older completion state.
- Migrates older sessions safely: a legacy `COMPLETE` state without a fingerprint requires one fresh handoff.
- After handoff completion, unresolved Creator candidates resume normally.
- Updates handoff UI wording to include Creator alongside Writer / Merge / monitoring.
