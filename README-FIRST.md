# SIMS Doctor Site Diagnosis v0.8.2 Patch

Base: v0.8.1

Fixes the post-generation checkpoint for individual precision diagnosis.

After a Doctor ZIP has been generated, `1. Site Diagnosisを進める` now recognizes
`PACKAGE_READY_FOR_DOCTOR` before considering new eligible articles. It displays the
existing filename and a Google Drive link and explicitly asks the user to give the ZIP
to SIMS Doctor.

`Doctorへ依頼しました` records the next checkpoint as `WAITING_DOCTOR_RESULT`.

This prevents accidental regeneration of the completed package and separates:
package generated -> Doctor requested -> Doctor result / SBM registration.
