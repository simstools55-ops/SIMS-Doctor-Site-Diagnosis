# RC5 Product UX Hotfix 2

## Purpose
Complete user-facing sheet cleanup after first Hotfix 1 UI verification.

## Changes
- Hide legacy English sheets when Japanese user-facing sheets already exist:
  - Site Diagnosis Candidates
  - Selected Treatment Cases
- Hide an unused blank default sheet:
  - シート1 / Sheet1
- Preserve legacy sheets and their data; do not delete them.
- Keep only the Japanese product sheets visible during normal operation:
  - 診断候補
  - 今回の診断対象

No diagnosis logic, scoring, request_id, Case Identity, or Case Package logic was changed.
