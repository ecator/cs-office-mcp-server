[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Workflow Status](https://github.com/ecator/cs-office-mcp-server/actions/workflows/build.yml/badge.svg)](https://github.com/ecator/cs-office-mcp-server/releases)

[🇨🇳中文](https://www.readme-i18n.com/zh/ecator/cs-office-mcp-server)
[🇯🇵日本語](https://www.readme-i18n.com/ja/ecator/cs-office-mcp-server)
[🇰🇷한국어](https://www.readme-i18n.com/ko/ecator/cs-office-mcp-server)
[🇩🇪Deutsch](https://www.readme-i18n.com/de/ecator/cs-office-mcp-server) 
[🇪🇸Español](https://www.readme-i18n.com/es/ecator/cs-office-mcp-server)
[🇫🇷français](https://www.readme-i18n.com/fr/ecator/cs-office-mcp-server)
[🇵🇹Português](https://www.readme-i18n.com/pt/ecator/cs-office-mcp-server)
[🇷🇺Русский](https://www.readme-i18n.com/ru/ecator/cs-office-mcp-server)

# Overview

The MCP Server for operating Office files such as Excel, Word, PowerPoint and Outlook.

**!!!Currently only supports Excel, Word and PowerPoint on Windows.**

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

Note that this is only supported on Windows with Office 2016 (64-bit version) or above installed!

# Tools
## Excel

### `excel_run_macro`
Run a macro of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `macroName*`: The name of macro.
- `macroParameters`: The parameters of macro. The maximum number is 30.
- `save`: Save the file after executing the macro.
- `password`: The password of the Excel file, if there is one.

### `excel_read_used_range`
Read the value of used range of cells from the specified worksheet.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_get_sheets`
Get all the sheet names of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_write`
Write data into a cell or a range of cells of the specified worksheet to an Excel file.
#### parameters
- `fullName*`: The full path of the Excel file. It will be created if not exist.
- `sheetName`: The sheet name of the Excel file. It will be created if not exist.
- `data`: The data that needs to be written in.
- `startColumn`: The first column as a letter where the data is written.(such as A)
- `startRow`: The first row number where the data is written.
- `password`: The password of the Excel file, if there is one.
- `forceOverwriteFile`: Force overwrite to create a new one when the file exists.
- `forceOverwriteSheet`: Force overwrite to create a new one when the sheet exists.

### `excel_rename_sheet`
Change the name of the sheet of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `oldSheetName*`: The old sheet name of the Excel file.
- `newSheetName*`: The new sheet name of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_get_tables`
Get all the table names of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_delete_sheet`
Delete the sheet of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_get_table_content`
Get the content of a table of the specified Excel file.
#### parameters
- `fullName*`: The full path of the Excel file.
- `tableName*`: The table name of the Excel file.
- `password`: The password of the Excel file, if there is one.

### `excel_find`
Find value from Excel files.
#### parameters
- `fullNameList*`: The list of full path of Excel files that need to be searched for.
- `searchValue*`: The value to be searched for which can use wildcard characters like `?`(any single character), `*`(any number of characters), `~` followed by `?`, `*`, or `~`(a question mark, asterisk, or tilde).
- `matchPart`: Match against any part of the search text when true. Match against the whole of the search text when false.
- `ignoreCase`: Ignoring lower case and upper case differences when true. Case insensitive when false
- `password`: The password of the Excel files, if there is one and all are the same.

### `excel_read`
Read the value of a cell or a range of cells from the specified worksheet.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `startColumn`: The first column as a letter.(such as A)
- `startRow`: The first row number.
- `endColumn`: The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn
- `endRow`: The last row number. If empty, then use xlDown relative to startRow
- `password`: The password of the Excel file, if there is one.

### `excel_clear`
Clear the value of a cell or a range of cells from the specified worksheet.
Clear the entire sheet if startColumn or startRow is null.
#### parameters
- `fullName*`: The full path of the Excel file.
- `sheetName*`: The sheet name of the Excel file.
- `startColumn`: The first column as a letter.(such as A)
- `startRow`: The first row number.
- `endColumn`: The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn
- `endRow`: The last row number. If empty, then use xlDown relative to startRow
- `password`: The password of the Excel file, if there is one.


## Word

### `word_run_macro`
Run a macro of the specified Word file.
#### parameters
- `fullName*`: The full path of the Word file.
- `macroName*`: The name of macro.
- `macroParameters`: The parameters of macro. The maximum number is 30.
- `save`: Save the file after executing the macro.
- `password`: The password of the Word file, if there is one.

### `word_clear`
Clear the whole content of the specified Word file.
#### parameters
- `fullName*`: The full path of the Word file.
- `password`: The password of the Word file, if there is one.

### `word_read`
Get the text content of the specified Word file.
#### parameters
- `fullName*`: The full path of the Word file.
- `fromPage`: The starting page number to read.
- `toPage`: The end page number to read. If it's empty, then read up to the last page.
- `password`: The password of the Word file, if there is one.

### `word_write`
Write data into an Word file.
#### parameters
- `fullName*`: The full path of the Word file. It will be created if not exist.
- `data`: The data that needs to be written in.
- `insertAfter`: Append to the end of the document when true. Append to the beginning of the document when false.
- `insertNewline`: Append a newline when writing to an existing file and the newline option is true. When data is appended to the end of the document, a newline character is added before the data. When data is prepended to the beginning of the document, a newline character is added after the data.
- `password`: The password of the Word file, if there is one.
- `forceOverwriteFile`: Force overwrite to create a new one when the file exists.

### `word_find`
Find value from Word files.
#### parameters
- `fullNameList*`: The list of full path of Word files that need to be searched for.
- `searchValue*`: The value to be searched for which can use wildcard characters like `?`(any single character), `*`(any number of characters), `\` followed by `?`, `*`, or `\`(a question mark, asterisk, or backslash).
- `matchPart`: Match against any part of part of a larger word when true. Match against the entire words of the search text when false.
- `ignoreCase`: Ignoring lower case and upper case differences when true. Case insensitive when false.
- `password`: The password of the Word files, if there is one and all are the same.

### `word_get_page_count`
Get all the number of the pages of the specified Word file.
#### parameters
- `fullName*`: The full path of the Word file.
- `password`: The password of the Word file, if there is one.

### `word_replace`
Replace value from Word files.
#### parameters
- `fullNameList*`: The list of full path of Word files that need to be searched for.
- `oldValue*`: The value to be searched for which can use wildcard characters like `?`(any single character), `*`(any number of characters), `\` followed by `?`, `*`, or `\`(a question mark, asterisk, or backslash).
- `newValue*`: The new replacement value.
- `matchPart`: Match against any part of part of a larger word when true. Match against the entire words of the search text when false.
- `ignoreCase`: Ignoring lower case and upper case differences when true. Case insensitive when false.
- `replaceAll`: Replace all matching items when true. Replace only the first matching item when false.
- `password`: The password of the Word files, if there is one and all are the same.

## PowerPoint

### `powerpoint_run_macro`
Run a macro of the specified PowerPoint file.
#### parameters
- `fullName*`: The full path of the PowerPoint file.
- `macroName*`: The name of macro.
- `macroParameters`: The parameters of macro. The maximum number is 30.
- `save`: Save the file after executing the macro.
- `password`: The password of the PowerPoint file, if there is one.

### `powerpoint_read`
Get the text content of the specified PowerPoint file.
#### parameters
- `fullName*`: The full path of the PowerPoint file.
- `fromSlide`: The starting slide number to read.
- `toSlide`: The end slide number to read. If it's empty, then read up to the last slide.
- `password`: The password of the PowerPoint file, if there is one.

### `powerpoint_find`
Find value from PowerPoint files.
#### parameters
- `fullNameList*`: The list of full path of PowerPoint files that need to be searched for.
- `searchValue*`: The value to be searched for.
- `matchPart`: Match against any part of part of a larger word when true. Match against the entire words of the search text when false.
- `ignoreCase`: Ignoring lower case and upper case differences when true. Case insensitive when false.
- `password`: The password of the PowerPoint files, if there is one and all are the same.

### `powerpoint_get_slide_count`
Get all the number of the slides of the specified PowerPoint file.
#### parameters
- `fullName*`: The full path of the PowerPoint file.
- `password`: The password of the PowerPoint file, if there is one.


## Outlook

Coming soon...
