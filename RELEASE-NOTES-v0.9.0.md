# Release Notes — SIMS Doctor Site Diagnosis v0.9.0

This MINOR release introduces the foundation for cross-collection diagnosis history.

Collector remains responsible only for supplying the latest Evidence Package.
Diagnosis now records a session summary keyed by site_id/host and keeps a case snapshot
when a diagnosis session is ended. Home can recognize and display the previous diagnosis
for the same site.

v0.9.0 intentionally does not yet calculate full per-article deltas such as:
recovered / continuously declining / newly declining. The new history store is the stable
foundation for that next phase.

Home's site-wide analysis comment area was also expanded so long diagnosis comments do not
appear clipped.
