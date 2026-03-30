# Payroll Editor Versions

This folder keeps versioned copies of the payroll upload editor in an English-named location.

## Versions

- `v1-original/payroll-upload-editor.original.html`
  - Byte-preserved copy of the current source file.
- `v3-excel-number-cleanup/payroll-upload-editor.excel-number-cleanup.html`
  - Export-focused version based on the original file.
  - Removes invisible whitespace noise from salary cells before Excel export.
  - Writes salary-area numeric-looking values as actual Excel numbers.
- `v4-styled-export/payroll-upload-editor.styled-export.html`
  - UI polished with Flowbite styling and updated typography.
  - Excel export uses ExcelJS for styled headers, alternating rows, comma-formatted numeric cells, and a totals row.

## Why The Delete Button Feels Broken

The delete button itself is wired correctly, but column selection is only attached to header cells.

- Selection handler: the click event is bound on `th` only.
- Guard clause: delete does nothing when `keptColumns.size === 0`.
- UX mismatch: both `th` and `td` show `cursor: pointer`, so body cells look clickable even though they do not select columns.

Relevant locations in the original file:

- `cursor: pointer` on both header/body cells
- `th.onclick = function() { toggleColumn(c, stepName); };`
- `if(keptColumns.size === 0) return alert(...)`

If you want, the next safe step is to make a UI-only version that improves the click affordance without changing the delete logic itself.
