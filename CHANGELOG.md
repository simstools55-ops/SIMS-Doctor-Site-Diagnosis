# CHANGELOG

## v0.8.2
- Fixed completed individual Doctor ZIP not being recognized by the workflow controller.
- Added PACKAGE_READY_FOR_DOCTOR workflow checkpoint.
- Added generated filename and Google Drive link.
- Added explicit `Doctorへ依頼しました` state transition.
- Prevented accidental ZIP regeneration after 9/9 completion.

## v0.8.1
- Fixed split preparation resume: 3/9 -> 6/9 -> 9/9.
