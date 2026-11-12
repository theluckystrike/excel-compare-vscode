'use strict';
// REACH1 (lane NEWSHIP): this listing's Pro tools for the toolbar row and the Explorer entries (src/reach.js).
module.exports = {
  "prefix": "excelCompare",
  "triesKey": "excelCompare.proTrial",
  "slot": "header.top",
  "tools": [
    {
      "id": "report",
      "title": "Export a Diff Report",
      "command": "excelCompare.exportReport",
      "rowOnly": true
    },
    {
      "id": "key",
      "title": "Match Rows by Key Columns",
      "command": "excelCompare.matchByKey",
      "rowOnly": true
    },
    {
      "id": "folders",
      "title": "Compare Two Folders of Workbooks",
      "command": "excelCompare.compareFolders",
      "folderMenu": true
    }
  ]
};
