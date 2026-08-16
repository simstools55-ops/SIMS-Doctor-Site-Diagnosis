# Release Notes — SIMS Doctor Site Diagnosis v0.7.4

## Fixed

The Site Treatment Plan previously remained visually unfinished after the user completed
registration in SIMS Blog Manager. The handoff state was stored internally, but the sheet
did not reflect that completion.

v0.7.4 makes that state visible:

- handed-off rows are grayed out;
- the State column displays `SBM引き渡し済み`;
- NO_ACTION rows display `SBM引き渡し済み（処置不要）`;
- unresolved NEEDS_EVIDENCE rows remain normal and active.

The completion view is tied to the existing SBM handoff fingerprint. If a new Creator case
or another treatment case changes the handoff payload, the old completion styling is no
longer considered current.

## Scope

No routing or treatment logic changes are included in this release.
