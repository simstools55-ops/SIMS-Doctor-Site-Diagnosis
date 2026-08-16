# SIMS Doctor Site Diagnosis v0.7.0

## New

- Added a Creator long-tail discovery route for YELLOW keyword clusters.
- Generates `SIMS_DOCTOR_LONGTAIL_DISCOVERY_REFERRAL_V1` for Doctor live-SERP research.
- Doctor can return up to five `+1 word` / narrowed-intent candidates with `SIMS_DOCTOR_LONGTAIL_DISCOVERY_RESULT_V1`.
- Added a `ロングテール探索` sheet and guided result-import dialog.
- A Doctor-approved CREATOR long-tail can be promoted to a Creator case and handed to SBM with a `creator_plan` and 30-day monitoring policy.

## Policy

- Do not create a second article for a keyword already owned by an existing article.
- Prefer narrowed intent modifiers and require SERP independence before Creator promotion.
- Zero discovered candidates is a valid result.
