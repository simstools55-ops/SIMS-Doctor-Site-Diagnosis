# RC9.3.2 Auto Article Identity

## Goal
Remove manual SBM Article Master CSV import from the normal precision-diagnosis workflow.

## Identity resolution
- Real SBM ArticleID available -> use it.
- Otherwise -> generate stable `REF-{hash}` from normalized article URL.
- Always store `canonical_article_url`.
- Mark `article_id_source` and `article_id_is_surrogate`.

## Content metadata
- Title: Article Master -> live page title
- Main query: Article Master -> top Collector query

## Important
A `REF-*` value is an internal Diagnosis identity, not an SBM ArticleID.
When returning treatment data to SBM, canonical URL must be used to resolve the real ArticleID.
