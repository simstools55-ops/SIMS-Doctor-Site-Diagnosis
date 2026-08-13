# Changelog

## 0.1.0-Sprint1-Hotfix3
- Fixed priority conversion so REVIEW ownership cannot be promoted to A1/A2/B by TVS alone.
- Priority order is now WAIT -> SBM -> REVIEW -> DOCTOR_OWNED TVS classification.
- No scoring, Article Universe, Ownership, or Recent Treatment Guard logic changed.

# Changelog

## 0.1.0-Sprint1-Hotfix2
- Added Article Universe gate before TVS scoring.
- Primary universe source is Collector RC5 page_query_top.
- Added Hatena/WordPress article-URL fallback.
- Site Analysis now reports Article Universe strategy/count.
- Prevents non-article / historical GSC URLs from being scored as current articles.

## 0.1.0-Sprint1-Hotfix1
- Fixed Apps Script syntax error in EvidenceModel.gs caused by mixing nullish coalescing (`??`) with logical OR (`||`) without parentheses.
- Replaced the mixed expression with an Apps Script-safe numeric fallback helper.


## 0.1.0-Sprint1
- Added Site Diagnosis configuration.
- Added Compact Evidence importer.
- Added SBM treatment history importer.
- Added URL normalization.
- Added Treatment Value Score V1 scoring.
- Added preliminary SBM_OWNED / DOCTOR_OWNED classification.
- Added Recent Treatment Guard.
- Added candidate output sheet.
