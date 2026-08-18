# Release Notes — SIMS Doctor Site Diagnosis v0.10.3

This PATCH release adds safe article-level routing for Site-wide Precision Diagnosis results after a real-site result returned mixed treatment routes inside a single cluster.

## Fixed

- Precision Result now accepts cluster-level `route_to: "MIXED"` only when every `articles[]` entry has an explicit final `route_to`.
- A MIXED cluster is expanded into concrete treatment cases before validation and registration.
- If all articles resolve to the same route, Diagnosis registers that concrete route instead of preserving MIXED.
- Article-level MERGE results using `merge_target_url` / `merge_target_title` are converted to the existing canonical Merge plan used by downstream handoff.
- Precision Result cluster identity now also accepts `site_diagnosis_case_id`, preserving the original Diagnosis case ID during split routing.
- Precision Doctor referral instructions now define the MIXED contract and prohibit MIXED without article-level routes.

## Safety

- `MIXED` itself is never stored as a final treatment route.
- MIXED clusters with missing article routes are rejected at VALIDATE.
- Existing allowed final routes remain `WRITER / MERGE / MONITOR / NO_ACTION / NEEDS_EVIDENCE`.
- Existing v0.10.2 complete-result and v0.10.1 large-text protections remain unchanged.

## Compatibility

- No `appsscript.json` change.
- Runtime replacement remains `Code.gs` only.
