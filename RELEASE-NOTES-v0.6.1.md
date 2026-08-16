# SIMS Doctor Site Diagnosis v0.6.1

Patch release for Creator-route upgrade-state migration and guided navigation.

## Fixes
- Creator candidates stored by older versions are automatically available to the v0.6.x Creator validation flow.
- A previously generated site-wide Precision package no longer blocks Creator validation. Its waiting state is preserved for later resumption.
- The Creator candidate view auto-generates the validation sheet when it does not yet exist.
- Home and guided navigation prioritize the Creator route when Creator candidates remain.

## Apps Script replacement
Replace `Code.gs`. No change is required to `appsscript.json`.
