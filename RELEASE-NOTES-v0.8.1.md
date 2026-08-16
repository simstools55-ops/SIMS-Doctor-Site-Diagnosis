# Release Notes — v0.8.1

This patch fixes the P1 split-execution defect found during the live v0.8.0 test.

Root cause:
`sdsdProceedIndividualPrecisionDiagnosis()` rebuilt the Treatment Batch at the beginning
of every run. `sdsdBuildTreatmentBatch()` clears and recreates the selected-case sheet,
so the three rows enriched in the previous run were lost.

Fix:
When an unfinished enrichment batch already exists, Diagnosis now resumes that batch
instead of rebuilding it. Already-ready rows are skipped by the existing enrichment logic.

Expected field-test sequence for the current nine-case batch:
3/9 prepared -> 6/9 prepared -> 9/9 prepared -> Doctor ZIP generated.
