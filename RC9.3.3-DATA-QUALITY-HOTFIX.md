# RC9.3.3 Data Quality Hotfix

## Problem found by Doctor
A Fitbit Air case contained:
- stale `article_evidence.title`
- `canonical_url` inconsistent with the article URL/current page

## Root cause
1. Article Master title was preferred over the live fetched page title.
2. HTML canonical was treated as the package's authoritative canonical URL.

## Fix
Title priority:
1. live fetched page title
2. Article Master title
3. URL-derived fallback

URL identity:
- `article_url` / `canonical_url`: Diagnosis target URL (authoritative)
- `observed_html_canonical_url`: HTML canonical observed during fetch
- `canonical_mismatch`: true/false
- mismatch details: `data_quality_flags`

This prevents stale or conflicting metadata from silently replacing case identity.
