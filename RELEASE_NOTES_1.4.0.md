# WiiBalanceScale v1.4.0

A stabilization-focused release that improves connection reliability, measurement clarity, profile-aware history, and export behavior.

## Highlights
- Improved Bluetooth reliability and scan-state handling, with clearer connection error mapping.
- Fixed Windows 11 quality indicator symbol rendering so quality stars display correctly.
- Expanded profile-aware behavior: optional profile height, profile-focused history, and profile-scoped exports.
- Improved session intelligence with plain-language summary/review text, comparisons, trend highlights, and posture/stability insights.
- Added QA/stability hardening for weak/early sessions, duplicate-save prevention, and safer history persistence.

## Included in this release
- Better handling around connection/session edge cases and null/transition states.
- Backward-compatible loading for existing `profiles.csv` and `session_history.csv` data.
- Cleaner fallback wording when a session does not have enough data for strong analysis.

## Version
- Application version: `1.4.0.0`
