# WiiBalanceScale v1.4.0 Manual Validation (Windows + Wii Balance Board)

Use this short pass before publishing the GitHub release.

## Test setup
- Windows 11 machine with Bluetooth enabled.
- Wii Balance Board with working batteries.
- Fresh launch of the v1.4.0 release build.

## Manual validation checklist
1. Launch app
   - Start the application and confirm main UI loads without errors.
2. Connect board
   - Press the board SYNC button and confirm the app connects.
3. Verify weight display
   - Stand on the board and confirm weight updates and stabilizes.
4. Verify quality indicator
   - Observe the quality stars/gauge and confirm symbol rendering is correct on Windows 11.
5. Verify profile with height
   - Create/select a profile with height set; confirm weight-vs-height guidance appears.
6. Verify profile without height
   - Create/select a profile with blank height; confirm no errors and sensible fallback text.
7. Verify session summary/review
   - Complete a session and confirm summary/review/advice text appears.
8. Verify export CSV
   - Export CSV and confirm file is created with selected-profile session rows.
9. Verify export JSON
   - Export JSON and confirm file is created with session summary fields.
10. Verify profile switching
    - Switch between profiles and confirm session/history context follows the selected profile.
11. Verify weak session behavior
    - Do a short/unstable session and confirm "not enough data" style messaging.
12. Verify trend/history behavior
    - Record multiple sessions and confirm trend/comparison/pattern text updates.
13. Verify shutdown/relaunch history persistence
    - Close app, relaunch, and confirm history/profiles are retained.

## Maintainer publish checklist
1. Build release binary (Release configuration).
2. Tag version `v1.4.0`.
3. Create GitHub Release titled `WiiBalanceScale v1.4.0`.
4. Attach release binary artifacts.
5. Confirm README release link path resolves correctly.
6. Paste final release notes body and publish.
