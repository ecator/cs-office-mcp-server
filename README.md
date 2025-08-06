[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Workflow Status](https://github.com/ecator/cs-office-mcp-server/actions/workflows/build.yml/badge.svg)](https://github.com/ecator/cs-office-mcp-server/releases)
# Overview

The MCP Server for operating Office files such as Excel,Word,PowerPoint,Outlook.

**!!!Currently only supports Excel on Windows.**

You must install Office 2016 and later versions to use this MCP server.

# Use

[Download the latest version](https://github.com/ecator/cs-office-mcp-server/releases) and extract it to any location.

Then add the following configuration to the MCP servers configuration.

```json
{
  "mcpServers": {
    "office": {
      "command": "DRIVER:\\PATH\\TO\\cs-office-mcp-server.exe",
      "args": [],
      "env": {}
    }
  }
}
```

Note that this is only supported on Windows that has Office 2016 and above installed in 64-bit version!

# Tools
## Excel

### excel_read_used_range
Read the value of used range of cells from the specified worksheet.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `password`: The password of the Excel file, if there is one.

### excel_get_sheets
Get all the sheet names of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `password`: The password of the Excel file, if there is one.

### excel_write
Write data into a cell or a range of cells of the specified worksheet.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `data*`: The data that needs to be written in.
- `startColumn`: The first column as a letter where the data is written.(such as A)
- `startRow`: The first row number where the data is written.
- `password`: The password of the Excel file, if there is one.

### excel_read
Read the value of a cell or a range of cells from the specified worksheet.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `startColumn`: The first column as a letter.(such as A)
- `startRow`: The first row number.
- `endColumn`: The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn
- `endRow`: The last row number. If empty, then use xlDown relative to startRow
- `password`: The password of the Excel file, if there is one.



TODO...

## Word

Coming soon...

## PowerPoint

Coming soon...

## Outlook

Coming soon...