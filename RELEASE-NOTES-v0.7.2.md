# Release Notes v0.7.2

Live testing found that a confirmed long-tail Creator case could not be handed to SBM while other YELLOW Creator candidates still remained. `▶ 次に進む` always returned to Creator candidate review first.

v0.7.2 changes the guided routing order: once any treatment is finalized, the current SBM handoff is offered first. This allows `ネイル 風水 就活` to continue through Diagnosis → SBM → Creator without requiring all remaining Creator candidates to be resolved first.

The patch also replaces the old session-wide `COMPLETE` flag semantics with a payload fingerprint check. If the treatment payload changes after an earlier handoff (for example, a long-tail Creator case is added), Diagnosis recognizes that a new SBM handoff is required. Once that exact payload is marked complete, the remaining Creator candidate workflow resumes.

The separate issue where a Creator case carries an existing related-article URL in `対象記事` is intentionally not changed in this patch. It remains under verification at the SBM handoff boundary.
