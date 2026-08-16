# Release Notes — v0.8.2

Live testing showed that reopening Site Diagnosis after 9/9 preparation and ZIP generation
incorrectly returned to `優先記事9件を詳しく診断します`.

v0.8.2 makes the generated ZIP an authoritative workflow checkpoint. The existing package
is shown instead of being regenerated. The user then explicitly records when the package
has been handed to SIMS Doctor.
