# SIMS Doctor Site Diagnosis v0.6.0

## Purpose

Adds the first complete Site Diagnosis Creator-validation route while keeping existing Writer/Merge/precision behavior intact.

## Regression case

Use the Creator candidate `財布 ベージュ 風水` as the first live validation case.

Expected route:

`Doctor primary site-wide result -> Site Diagnosis Creator candidate check -> Doctor SERP check -> Site Diagnosis -> SBM -> Creator -> SBM monitoring`

## Important policy

- Avoid clear same-intent cannibalization.
- Do not block a long-tail article merely because related articles exist.
- Allow calculated experimentation for GREEN/YELLOW cases.
- Judge the published article from real GSC data in SBM at roughly 30 days; extend monitoring when data is insufficient.

## Files to replace

- `Code.gs`

Repository/release metadata updated together:

- `VERSION`
- `README-FIRST.md`
- `CHANGELOG.md`
- `RELEASE-NOTES-v0.6.0.md` (new)

No change intended:

- `appsscript.json`
- `SITE-WIDE-RESULT-CONTRACT-V1.md`
