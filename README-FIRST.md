# SIMS Doctor Site Diagnosis v0.7.3

PATCH for Creator new-article identity semantics at the Site Diagnosis -> SBM handoff boundary.

## Apps Script replacement
- Replace: `Code.gs`
- No other Apps Script files are required.

## Main fix
A Creator long-tail case no longer presents an existing related article as the actual `対象記事` in Site Diagnosis. The treatment plan now shows the true target as a not-yet-published new article and its candidate keyword.

For backward compatibility with the current SBM identity verification, existing related articles are still carried only as reference/bridge identity data. v0.7.3 marks that role explicitly with `REFERENCE_ONLY`, adds `new_article_target`, `reference_articles`, and `article_identity_semantics`, and keeps the existing bridge payload usable until SBM consumes the new semantics directly.

## Version consistency
`Code.gs`, `VERSION`, README, CHANGELOG, and release notes are aligned to v0.7.3.
