# SIMS Doctor Site Diagnosis v0.7.2

PATCH for SBM handoff priority in the Creator long-tail workflow.

## Apps Script replacement
- Replace: `Code.gs`
- No other Apps Script files are required for this patch.

## Main fix
When at least one Writer / Merge / Creator treatment is finalized, `▶ 次に進む` now prioritizes the SBM handoff before continuing unresolved Creator candidates. After the current handoff is completed, remaining YELLOW Creator candidates can be resumed.

v0.7.2 also fingerprints the completed SBM payload. This makes an older session with a newly-added Creator case automatically require a fresh handoff, while preventing the same unchanged payload from being requested repeatedly.
