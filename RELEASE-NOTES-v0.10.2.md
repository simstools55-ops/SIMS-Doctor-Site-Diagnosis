# Release Notes — SIMS Doctor Site Diagnosis v0.10.2

This PATCH release fixes an incomplete-result issue discovered in a 345-case real-site Site-wide Diagnosis run.

## Fixed

- Site-wide Doctor referrals now require `SIMS_DOCTOR_SITE_WIDE_RESULT_V1` contract version 1.1 with a complete `diagnosis_cases[]` array.
- Doctor is explicitly forbidden from returning only representative/top-priority clusters, omitting the remainder, or promising a separate file later.
- The result must include `result_complete: true`, `returned_diagnosis_case_count`, and `omitted_diagnosis_case_count: 0`.
- Diagnosis rejects a result when `cluster_count` does not match `diagnosis_cases.length`.
- Diagnosis verifies full source-case coverage by comparing every Package input `case_id` with the union of `diagnosis_cases[].absorbed_source_case_ids`.
- Diagnosis rejects results routed to anything other than `SIMS_DOCTOR_SITE_DIAGNOSIS`.
- `manifest.json` now advertises the required response contract and source-case count.

## Safety

- Incomplete Doctor output is rejected at VALIDATE before replacing the stored Site-wide result.
- Existing v0.10.1 large-text chunking remains unchanged.
- Legacy 1.0 complete results remain accepted unless they explicitly indicate omitted/remaining cases.

## Compatibility

- No `appsscript.json` change.
- Runtime replacement remains `Code.gs` only.
