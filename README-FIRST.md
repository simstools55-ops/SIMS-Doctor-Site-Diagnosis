# SIMS Doctor Site Diagnosis v0.8.1 Patch

Base: v0.8.0

## P1 fix
Individual precision diagnosis batch preparation now resumes correctly.

Previously every click rebuilt the selected treatment batch, clearing the three already
prepared rows. This caused the progress to remain at 3/9 indefinitely.

v0.8.1 detects an unfinished selected-case batch and reuses it. A 9-case batch therefore
progresses 3/9 -> 6/9 -> 9/9 instead of restarting at 3/9.

## UI improvements from the v0.8.0 field test
- Resume dialog shows prepared / total / remaining counts.
- The next button says how many remaining cases will be prepared.
- Old `次に進む（Diagnosisに任せる）` wording is replaced by `1. Site Diagnosisを進める`.
- After package creation, the UI explicitly instructs the user to give the ZIP to SIMS Doctor.
- The progress dialog explains the relationship between all additional-Evidence cases and the current high-priority individual batch.

No Doctor/SBM contract changes are included.
