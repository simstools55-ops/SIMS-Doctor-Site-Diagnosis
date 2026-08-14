# RC5 Product UX Hotfix 5

## Problem
Menu 7 exceeded the Google Apps Script maximum execution time while preparing 12 Doctor Case Packages.

## Fix
- Case Enrichment is resumable.
- A single menu-7 execution prepares at most 3 unfinished cases.
- Existing READY cases are skipped and never fetched again.
- Progress is saved in `Referral Status` and `Case Package Status`.
- The next menu-7 execution resumes from unfinished cases.
- ZIP export begins only after every selected case is READY.
- Existing Treatment Batch and imported Article Master are reused.
- Menus 1-6 do not need to be rerun.

## Expected 12-case flow
Typical sequence:
- Run 1: 3/12 ready
- Run 2: 6/12 ready
- Run 3: 9/12 ready
- Run 4: 12/12 ready, then ZIP export

The exact ready count may differ if some rows were already completed before a timeout.

## Protected behavior
No changes to:
- diagnosis scoring
- Treatment Batch selection
- Case Identity
- request_id
- referral contract
- article/query evidence requirements
