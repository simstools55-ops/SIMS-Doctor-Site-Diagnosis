# SIMS_DOCTOR_SITE_WIDE_RESULT_V1

Required top-level fields:
- format = `SIMS_DOCTOR_SITE_WIDE_RESULT_V1`
- contract_version
- site
- site_diagnosis_batch_id
- overall_diagnosis
- diagnosis_cases[]

Each diagnosis case:
- diagnosis_case_id
- diagnosis_theme
- diagnosis_type
- absorbed_source_case_ids[]
- target_articles[]
- doctor_decision
- confidence
- site_impact
- treatment_strategy
- route_to
- eventual_route (optional)
- reason
- additional_evidence_needed[]

Allowed `route_to`:
- WRITER
- MERGE
- CREATOR
- MONITOR
- NO_ACTION
- NEEDS_EVIDENCE

Safety rule:
A Diagnosis `新規記事機会` does not directly authorize Creator.
Doctor must explicitly return `route_to = CREATOR`.
