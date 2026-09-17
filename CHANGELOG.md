# Changelog

## 2.5.13

- Disabled `Tab` and `Shift+Tab` while the LEFT FILES Treeview has keyboard focus.
- Prevents default Tk focus traversal from jumping to toolbar buttons with no obvious visual destination.
- Pane switching remains explicit via `Ctrl+Left` / `Ctrl+Right` and `Alt+1` / `Alt+2`.
- RIGHT WORKBOOK keeps normal Excel-style `Tab` / `Shift+Tab` cell navigation.
- Added regression coverage for the LEFT file-list Tab policy.

## 2.5.12

- Fixed stale RIGHT-pane workbook content when the LEFT selection moves from an Excel workbook to a folder or non-Excel file.
- Cancels pending/deferred workbook loads and invalidates in-flight background loaders before clearing the viewer.
- Restores the blank Workbook placeholder after clearing so no previous Excel image remains visible.
- Protects unsaved workbooks: Save / Discard / Cancel is respected before clearing the currently edited workbook.
- Added regression coverage for selection-to-viewer action decisions.

## 2.5.11

- LEFT Search box: `Up` / `Down` now immediately applies any pending filter, exits the search field, focuses the filtered file list, and continues selection navigation.
- `Down` enters the filtered list at the first result; `Up` enters it at the last result, avoiding stale pre-search selections.
- `Enter` now also flushes the current debounced search text before selecting the first result.
- Added regression coverage for search-to-file arrow navigation.

## 2.5.10

- Final GitHub-readiness audit and packaging cleanup.
- Unified app/viewer version reporting with the package `__version__`.
- Removed stale hard-coded installer version labels.
- Installer recreates incompatible pre-existing private venvs.
- Uninstaller removes stale Excel `OpenWithProgids` registrations.
- Added Windows GitHub Actions CI for Python 3.11 / 3.12 / 3.13.
- Added clean release ZIP + SHA256 packaging tool and a version-consistency regression test.

## 2.5.9

- Fixed LEFT-pane RIGHT-DRAG selection when the gesture starts on an already-selected file.
- The starting row remains selected and every crossed row is added to the selection.
- Existing selections are not accidentally toggled off.
- Fast-drag gap filling and edge auto-scroll remain enabled.

## 2.5.8

- Added mDIR-style right mouse drag selection in the LEFT file list.
- Fast drags include intermediate rows and edge dragging auto-scrolls.
- Workbook auto-loading pauses during the gesture and resumes after release.

## 2.5.7

- Added standard Ctrl-click / Shift-click multi-selection in the LEFT file list.
- Del moves selected files/folders to the Windows Recycle Bin after confirmation.
- Unsaved open workbooks are protected before deletion.

## 2.5.6

- Enter on a LEFT-pane file opens it with the Windows-associated application.
- Enter on a folder navigates into that folder.

## 2.5.5

- Added persistent recent-directory history with a compact dropdown beside the path bar.

## 2.5.4

- Fixed Ctrl+F filename search so ordinary number keys are not misread as Alt+1 / Alt+2 pane shortcuts on Windows.

## 2.5.3

- Added live LEFT-pane filename filtering, Ctrl+F focus, Esc clear, and cached directory filtering.

## 2.5.2

- Split the RIGHT WORKBOOK header into a button row and a separate filename/status row.

## 2.5.1

- Extended RIGHT WORKBOOK scroll extents to the full bottom/right edge of embedded images.

## 2.5.0

- Added Bright / Dark / System themes with persistent preference and Windows system-theme following.

## 2.4.1

- Softened the LEFT/RIGHT pane divider while preserving resize behavior.

## 2.4.0

- Added the mDIR-compatible Link Manager and full Folder / File / Program / Web / Action / Command shortcut model.

## 2.3.1

- Replaced breadcrumb buttons with the continuous green mDIR-style clickable path bar.

## 2.3.0

- Added mDIR folder shortcut import and editable quick links.

## 2.2.0

- Improved Excel layout fidelity: hidden rows/columns, row heights, column widths, gridline state, fonts, fills, borders, alignment, and image anchors.

## 2.1.4

- Hardened Ctrl+Left / Ctrl+Right and Alt+1 / Alt+2 pane switching on Windows/Tk.

## 2.1.3

- Made Ctrl+Left / Ctrl+Right the primary pane-navigation shortcuts while retaining Alt+1 / Alt+2 and F6.

## 2.1.2

- Added dedicated pane-navigation shortcuts and returned Tab / Shift+Tab to cell navigation.

## 2.1.1

- Fixed LEFT pane collapse at startup and stabilized two-pane geometry.

## 2.1.0

- Added mouse/keyboard pane focus, LEFT file arrow navigation, RIGHT cell arrow navigation, and full worksheet tab display.

## 2.0.0

- Introduced the integrated LEFT file browser + RIGHT full-workbook viewing/editing workspace.
