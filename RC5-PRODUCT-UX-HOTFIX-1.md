# RC5 Product UX Hotfix 1

## Fixed during first UI test
- Fixed Initial Setup failure by routing setup through a self-contained product sheet initializer.
- Renamed user-facing sheets:
  - `Site Diagnosis Candidates` -> `診断候補`
  - `Selected Treatment Cases` -> `今回の診断対象`
- Added five Japanese user-facing columns to each visible sheet.
- Preserved technical columns on the same sheets but hid them from normal users.
- Hide an unused blank default `シート1` / `Sheet1` sheet.
- Diagnosis logic, Case Identity, request_id and scoring rules are unchanged.

## User-facing columns
### 診断候補
- 順位
- 記事URL
- 診断優先度
- 選定理由
- 現在の状態

### 今回の診断対象
- No.
- 記事URL
- 診断優先度
- 選定理由
- 状態
