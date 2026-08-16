# SIMS Doctor Site Diagnosis v0.8.0 Patch

Base: v0.7.4

## Major workflow change

Site Diagnosis now has two large user-facing flows:

1. Existing article diagnosis / treatment
2. New article opportunity / Creator

After the Collector Evidence Package is loaded, Diagnosis analyzes the whole site first.
It compares both flows and recommends which one should be handled first.

Normal operation starts from:

`1. Site Diagnosisを進める`

A modal workflow controller shows the current route, current step, required input/output,
and the next action. The user can close the dialog at any time. The selected route and
workflow state remain saved, so opening the same command resumes from the current step.

The user can switch to the other large flow without discarding completed work.

## Priority principle

Existing-article issues receive higher priority when there are strong A1/A2 candidates or
clear severe / traffic / ranking declines. Creator receives priority when new-article
opportunities and content gaps outweigh urgent existing-article work.

The recommendation is advisory. The user can choose the other route.

## Compatibility

Underlying Doctor, Evidence, SBM handoff, Writer, Merge and Creator contracts are unchanged.
