# Release Notes v0.7.3

Live Creator-route testing exposed an identity-semantics problem: a newly-created Creator case inherited the related existing article list from its parent validation case. Site Diagnosis therefore displayed an existing article under `対象記事`, even though the actual treatment target was a brand-new, unpublished article.

v0.7.3 separates those two meanings without breaking the current SBM bridge.

- The Creator case now has an explicit `new_article_target` with the candidate keyword and `NOT_PUBLISHED` status.
- Existing articles used for cannibalization checks, internal-link design, and legacy SBM identity verification are marked `REFERENCE_ONLY` and exposed as `reference_articles`.
- `サイト治療計画` shows `新規記事（未発行） / キーワード: ...` for Creator cases instead of showing the related existing URL as the treatment target.
- The SBM handoff payload includes `article_identity_semantics` so the next SBM-side patch can distinguish the new article target from the compatibility identity article.
- The legacy `articles` identity payload is retained for now because current SBM registration validates Site Diagnosis cases against an existing ArticleID and URL.

This patch intentionally avoids removing the compatibility identity article until SBM is updated to accept Creator cases without pretending that the reference article is the treatment target.
