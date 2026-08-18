# Release Notes — SIMS Doctor Site Diagnosis v0.10.1

This PATCH release fixes issues found during a large real-site Site Diagnosis run.

## Fixed

- Doctor responses over the Google Sheets 50,000-character-per-cell limit are now stored automatically in 40,000-character chunks and reconstructed before JSON extraction. The user still pastes the Doctor response only once.
- Site-wide batch IDs are cleared when a diagnosis session ends and when a new Evidence Package starts a new session, preventing a batch ID from being reused across different sites.
- Site identity in the site-wide Doctor package now falls back to the active Diagnosis session when Article Master metadata is incomplete.
- Site-wide Doctor referral instructions now state explicitly that the response returns to SIMS Doctor Site Diagnosis, not directly to SBM, and require the machine-readable JSON body to be included.
- User-facing site-wide package names now show the route, readable blog name, purpose, and timestamp, e.g. `SIMS-Diagnosis-to-Doctor-ガジェット探検記-Site-Wide-Diagnosis-YYYYMMDD-HHMMSS.zip`.
- The session-end menu is renamed to make the next-blog workflow clearer.

## Compatibility

- No `appsscript.json` change.
- Existing short Doctor responses and manual import-sheet workflows remain supported.
- Existing diagnosis history behavior is unchanged.
