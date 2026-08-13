# Changelog

## 0.3.0-Sprint3
- Added site-size treatment capacity policy.
- Added automatic A1/A2 treatment batch selection.
- Selection uses standardMax as an upper bound, never as a quota.
- B/REVIEW/WAIT/SBM/PROTECTED cases are not used to fill the batch.
- Added Selected Treatment Cases sheet.
- Added draft Individual Doctor referral JSON per selected case.
- Added Final Guard recheck against latest SBM treatment history.
- Article body/ArticleID enrichment is intentionally deferred to the next SBM integration step.

## 0.2.2-Sprint2-Weekly-Aggregation-Hotfix
- Fixed weekly trend analysis to aggregate duplicate/raw URL variants by normalized URL and week.
- Weekly clicks/impressions are summed.
- Weekly average position is recalculated using impression-weighted average.
- Week keys are normalized before grouping.
- No TVS, Ownership, Recent Treatment Guard, Risk, or Priority rules changed.

## 0.2.0-Sprint2
- Added weekly trend classification.
- Added Evidence Confidence.
- Added Treatment Risk.
- Added external-factor flags from page query evidence.
- Added PROTECTED for growth cases.
- Added DOCTOR_REVIEW for unresolved high-value/review cases.
- Expanded candidate output columns.

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
