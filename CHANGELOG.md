# Changelog

## 0.1.0

- First release.
- Compare two Excel or CSV files (.xlsx, .xlsm, .xls, .ods, .csv, .tsv): Compare Selected Excel Files on two selected files, Compare with... on one, and a command.
- Compare a file with its Git HEAD version.
- Rows are lined up (inserted and removed rows do not shift the rest), columns are matched by their header, sheets by name and then by position.
- Changed, added and removed cells and rows in colour, next and previous change, counts per sheet, values or formulas.
- Grid that only draws the rows on screen, for sheets with hundreds of thousands of rows. Limits on file size, inflate size, rows and columns, each with a note when it applies.
- Zovo Pro tools (3 free tries each): export a diff report (XLSX with coloured cells, HTML or CSV), match rows by key columns with moved rows, compare two folders of workbooks.
- A Pro tool uses no free try when nothing differs: no report for two files that are the same, no try for a key match with no change, or for two folders whose files are all the same.
- No network and no telemetry.
