# WiiBalanceScale 1.4.0 (Draft)

Stabilization and release-readiness update focused on QA hardening, edge-case handling, and clearer plain-language messaging.

## Highlights since 1.3.1
- Added runtime hardening for session finalization, repeated-save deduplication, and safer history writes.
- Improved backward compatibility handling for `profiles.csv` and legacy session history shapes.
- Polished edge-case messaging for weak/early sessions and ambiguous trend/pattern outcomes.
- Improved profile handling by allowing optional (blank) height entries.
- Updated exports to be profile-focused (selected profile only) to avoid cross-profile mixed output.
- Reduced risk of UI text overflow by trimming long summary/review labels for readability.

## QA / Stability hardening
- Added defensive guard for connection tick handling when the balance-board object is unexpectedly null.
- Added per-profile session end tracking to prevent duplicate history records on repeated saves and profile switch/save cycles.
- Switched session history persistence to a temporary-file write then move strategy to reduce corruption risk on interrupted saves.
- Normalized key serialized values (non-negative sample count/height/time and bounded stability score).

## Backward compatibility
- `profiles.csv` now tolerates headers/comments and quoted CSV values while still accepting old formats.
- `session_history.csv` loading remains compatible with older column sets (existing fallback behavior retained).

## Product polish
- Clarified profile height input behavior with explicit support for blank height.
- Improved trend text fallback when there is no clear directional signal.
- Clear-session now resets advice text immediately to avoid stale guidance.

## Version
- Application version: `1.4.0.0`
