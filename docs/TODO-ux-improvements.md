# TODO: UX Improvements — Loading Indicator + Convert to Values

**Branch:** `feature/ux-loading-convert`
**Created:** 2026-03-06

## Phase 1: Loading Indicator

Replace `#N/A` shown during async API calls with "Loading..." message.

- [x] Add `LoadingObservable` class (`IExcelObservable`) to `UseAiFunction.cs`
- [x] Add `ActionDisposable` helper class
- [x] Switch `CallLlm()` from `ExcelAsyncUtil.Run()` to `ExcelAsyncUtil.Observe()`
- [ ] Verify spill arrays (USEAI) work through Observe pattern *(manual test in Excel)*
- [x] Build and verify — **0 warnings, 0 errors**

## Phase 2: Convert to Values (Share Workbook)

Add buttons to replace USEAI formulas with their computed values for sharing.

- [x] Add "Tools" tab to `SettingsForm.cs`
- [x] Implement "Convert Selected Cells" (replaces selection formulas with values)
- [x] Implement "Convert All USEAI Cells" (finds and converts all USEAI formulas in workbook)
- [x] Add "Convert to Values" ribbon button in `RibbonController.cs`
- [x] Build and verify — **0 warnings, 0 errors**

## Phase 3: Documentation

- [x] Update `docs/SRS.md` — §3.1.5 (loading), §3.4.4 (Tools tab), §3.8 (Convert to Values), §5.3
- [x] Update `CLAUDE.md` — architecture + service descriptions
