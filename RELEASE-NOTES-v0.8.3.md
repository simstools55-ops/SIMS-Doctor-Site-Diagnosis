# Release Notes — v0.8.3

Live testing showed that Doctor returned valid common Writer scopes at the root
`workflow_handoff`, but Diagnosis v0.8.2 discarded them when converting Precision Result
clusters. The generated SBM JSON therefore contained empty `allowed_scope` arrays.

v0.8.3 preserves the scope contract and blocks unsafe handoff generation if a Writer
cluster still has no allowed treatment scope.
