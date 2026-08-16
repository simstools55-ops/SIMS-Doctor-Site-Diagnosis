# RC8 Diagnosis Session Lifecycle

## Purpose

Site Diagnosis is a shared diagnostic workspace, not a permanent multi-site database.
RC8 prevents Evidence and diagnosis data from different sites being mixed.

## User workflow

1. Import one site's Evidence Package.
2. Run diagnosis and hand required cases to Doctor / SBM.
3. Before diagnosing another site, use `診断セッション > 現在の診断を終了`.
4. Import the next site's latest Evidence Package.

## What is cleared when a session ends

- Collector Evidence working copies
- Site diagnosis summary / candidates
- Selected precision-diagnosis cases
- Site-wide opportunities / Doctor cases / treatment plan
- Diagnosis-side copies of SBM history and Article Master
- Imported Doctor site-wide result and Merge referral working data

The original SBM data is not modified.

## Protection

A new Evidence Package cannot be imported while a diagnosis session is active.
The session-end dialog shows remaining work counts before clearing data.
