# xExcel Viewer 2.5.13 Release Notes

## LEFT FILES Tab behavior is now explicit

When the LEFT file list has keyboard focus, `Tab` and `Shift+Tab` are intentionally blocked. This prevents Windows/Tk focus traversal from moving to toolbar buttons where the current focus can be difficult to see.

- LEFT FILES: `Tab` / `Shift+Tab` do nothing and focus remains on the file list.
- Move to RIGHT WORKBOOK with `Ctrl+Right` or `Alt+2`.
- Return to LEFT FILES with `Ctrl+Left` or `Alt+1`.
- RIGHT WORKBOOK: `Tab` / `Shift+Tab` continue to move to the next / previous Excel cell.
- Existing search, multi-select, right-drag selection, recycle-bin deletion, stale-view clearing, themes, history, links, full workbook rendering, editing, and image handling remain unchanged.
