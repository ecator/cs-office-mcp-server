using ModelContextProtocol;
using ModelContextProtocol.Server;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class ExcelTools

{

    [McpServerTool(Name = "excel_get_sheets"), Description("Get all the sheet names of the specified Excel file.")]
    public static string GetSheets([Description("The full path of the Excel file.")] string fullName, [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        var data = new StringBuilder();
        var count = 0;
        data.AppendLine();
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);

            foreach (Excel.Worksheet sheet in session.GetSheets(wk))
            {
                count++;
                data.AppendLine($"{count}. {sheet.Name}");
            }
        }
        data.Insert(0, $"Total `{count}` sheets in the Excel file `{fullName}`:");
        return data.ToString();
    }

    [McpServerTool(Name = "excel_get_tables"), Description("Get all the table names of the specified Excel file.")]
    public static string GetTables([Description("The full path of the Excel file.")] string fullName, [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        var data = new StringBuilder();
        var count = 0;
        data.AppendLine();
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);

            foreach (Excel.ListObject table in session.GetTables(wk))
            {
                count++;
                data.AppendLine($"{count}. {table.Name}");
            }
        }
        data.Insert(0, $"Total `{count}` tables in the Excel file `{fullName}`:");
        return data.ToString();
    }

    [McpServerTool(Name = "excel_get_table_content"), Description("Get the content of a table of the specified Excel file.")]
    public static string GetTableContent([Description("The full path of the Excel file.")] string fullName
        , [Description("The table name of the Excel file.")] string tableName
        , [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        Excel.ListObject table = null;
        Excel.Range range;
        Excel.Range cells;
        StringBuilder data = new StringBuilder();
        string address;
        string sheetName;
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);

            foreach (Excel.ListObject t in session.GetTables(wk))
            {
                if(t.Name == tableName)
                {
                    table = t;
                    break;
                }
            }
            if(table == null)
            {
                throw new McpException($"The table {tableName} does not exist in the Excel file {fullName}.");
            }
            range = table.Range;
            session.RegisterComObject(range);
            var rows = range.Rows;
            session.RegisterComObject(rows);
            var cols = range.Columns;
            session.RegisterComObject(cols);
            cells = range.Cells;
            session.RegisterComObject(cells);
            Excel.Worksheet sh = table.Parent as Excel.Worksheet;
            session.RegisterComObject(sh);
            sheetName = sh.Name;
            address = range.Address[false,false];
            data.AppendLine($"Table {tableName} is in sheet {sheetName}, address {address}, total {rows.Count - 1} data rows and {cols.Count} columns.");
            for (int i = 1; i <= rows.Count; i++)
            {
                var lines = new string[cols.Count];
                if (i == 2)
                {
                    Array.Fill(lines, "---");
                    data.AppendLine(string.Join("|", lines));
                }
                for (int j = 1; j <= cols.Count; j++)
                {
                    Excel.Range cell = cells[i, j];
                    session.RegisterComObject(cell);
                    string s = Convert.ToString(cell.Value) ?? "";
                    s = session.EscapeMarkdownTableValue(s);
                    lines[j - 1] = s;
                }
                data.AppendLine(string.Join("|", lines));
            }

        }

        return data.ToString();
    }

    [McpServerTool(Name = "excel_read"), Description("Read the value of a cell or a range of cells from the specified worksheet.")]
    public static Dictionary<string, object> Read([Description("The full path of the Excel file.")] string fullName
        , [Description("The sheet name of the Excel file.")] string sheetName
        , [Description("The first column as a letter.(such as A)")] string startColumn = "A"
        , [Description("The first row number.")] int startRow = 1
        , [Description("The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn")] string? endColumn = null
        , [Description("The last row number. If empty, then use xlDown relative to startRow")] int? endRow = null
        , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        var values = new Dictionary<string, object>();

        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);
            var sh = session.GetSheet(wk, sheetName);
            var range = session.GetRange(sh, startColumn, startRow, endColumn, endRow);
            foreach (Excel.Range r in range)
            {
                session.RegisterComObject(r);
                values[r.Address.Replace("$", "")] = r.Value;
            }
        }

        return values;
    }
    [McpServerTool(Name = "excel_read_used_range"), Description("Read the value of used range of cells from the specified worksheet.")]
    public static Dictionary<string, object> ReadUsedRange([Description("The full path of the Excel file.")] string fullName
        , [Description("The sheet name of the Excel file.")] string sheetName
        , [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        var values = new Dictionary<string, object>();

        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);
            var sh = session.GetSheet(wk, sheetName);
            var range = sh.UsedRange;
            session.RegisterComObject(range);
            foreach (Excel.Range r in range)
            {
                session.RegisterComObject(r);
                values[r.Address.Replace("$", "")] = r.Value;
            }
        }

        return values;
    }

    [McpServerTool(Name = "excel_find"), Description("Find value from Excel files.")]
    public static string Find([Description("The list of full path of Excel files that need to be searched for.")] string[] fullNameList
    , [Description("The value to be searched for which can use wildcard characters like ?(any single character), *(any number of characters), ~followed by ?, *, or ~(a question mark, asterisk, or tilde).")] string searchValue
    , [Description("Match against any part of the search text when true. Match against the whole of the search text when false.")] bool matchPart = true
    , [Description("Ignoring lower case and upper case differences when true. Case insensitive when false.")] bool ignoreCase = true
    , [Description("The password of the Excel files, if there is one and all are the same.")] string? password = null)
    {
        var data = new StringBuilder();
        var foundData = new StringBuilder();
        var line = new string[3];
        var totalCount = 0;
        var count = 0;
        var lookAt = Excel.XlLookAt.xlPart;
        if(!matchPart)
        {
            lookAt = Excel.XlLookAt.xlWhole;
        }
        if (fullNameList == null || fullNameList.Length == 0)
        {
            throw new McpException("The full path list of the Excel file cannot be empty or null.");
        }
        data.AppendLine();
        data.AppendLine();
        using (var session = new ExcelSession())
        {
            foreach (var fullName in fullNameList)
            {
                var wk = session.OpenWorkbook(fullName, true, password);
                var shs = session.GetSheets(wk);
                count = 0;
                foundData.Clear();
                foundData.AppendLine();
                foreach (Excel.Worksheet sh in shs)
                {
                    var cells = sh.Cells;
                    session.RegisterComObject(cells);
                    var found = cells.Find(searchValue, Type.Missing, Excel.XlFindLookIn.xlValues, lookAt, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, !ignoreCase, Type.Missing, Type.Missing);
                    if (found != null)
                    {
                        session.RegisterComObject(found);
                        var firstAddress = found.Address;
                        if(count == 0)
                        {
                            foundData.AppendLine($"Sheet|Address|Value");
                            foundData.AppendLine($"---|---|---");
                        }
                        do
                        {
                            totalCount++;
                            count++;
                            line[0] = sh.Name;
                            line[1] = found.Address[false, false];
                            line[2] = session.EscapeMarkdownTableValue(Convert.ToString(found.Value));
                            foundData.AppendLine(string.Join("|", line));
                            found = cells.FindNext(found);
                            if (found != null)
                            {
                                session.RegisterComObject(found);
                            }
                        } while (found != null && found.Address != firstAddress);
                    }
                }
                if(count > 0)
                {
                    foundData.Insert(0, $"`{count}` results in `{fullName}`:");
                    data.AppendLine(foundData.ToString());
                }
            }

        }
        data.Insert(0, $"Found a total of `{totalCount}` results for `{searchValue}` in all files.");
        return data.ToString();
    }

    [McpServerTool(Name = "excel_clear"), Description("Clear the value of a cell or a range of cells from the specified worksheet.\nClear the entire sheet if startColumn or startRow is null.")]
    public static string Clear([Description("The full path of the Excel file.")] string fullName
    , [Description("The sheet name of the Excel file.")] string sheetName
    , [Description("The first column as a letter.(such as A)")] string? startColumn = null
    , [Description("The first row number.")] int? startRow = null
    , [Description("The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn")] string? endColumn = null
    , [Description("The last row number. If empty, then use xlDown relative to startRow")] int? endRow = null
    , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, false, password);
            var sh = session.GetSheet(wk, sheetName);
            Excel.Range range;
            if (string.IsNullOrEmpty(startColumn) || startRow == null)
            {
                range = sh.Cells;
                response = "All cells cleared.";
            }
            else
            {
                range = session.GetRange(sh, startColumn, startRow.Value, endColumn, endRow);
                response = $"{range.Address} cleared.";
            }
            session.RegisterComObject(range);
            range.Clear();
            wk.Save();
            wk.Close();
        }

        return response;
    }
    [McpServerTool(Name = "excel_rename_sheet"), Description("Change the name of the sheet of the specified Excel file.")]
    public static string RenameSheet([Description("The full path of the Excel file.")] string fullName
    , [Description("The old sheet name of the Excel file.")] string oldSheetName
    , [Description("The new sheet name of the Excel file.")] string newSheetName
    , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        var response = "";
        if (string.IsNullOrEmpty(oldSheetName))
        {
            throw new McpException("The old sheet name cannot be empty or null.");
        }
        if (string.IsNullOrEmpty(newSheetName))
        {
            throw new McpException("The new sheet name cannot be empty or null.");
        }
        if (oldSheetName.Equals(newSheetName))
        {
            throw new McpException("The new sheet name cannot be the same as old sheet name.");
        }
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, false, password);
            var sh = session.GetSheet(wk, oldSheetName);
            sh.Name = newSheetName;
            wk.Save();
            wk.Close();
        }
        response = $"{oldSheetName} has been changed to {newSheetName}.";
        return response;
    }
    
    [McpServerTool(Name = "excel_delete_sheet"), Description("Delete the sheet of the specified Excel file.")]
    public static string DeleteSheet([Description("The full path of the Excel file.")] string fullName
    , [Description("The sheet name of the Excel file.")] string sheetName
    , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, false, password);
            var sh = session.GetSheet(wk, sheetName);
            sh.Delete();
            wk.Save();
            wk.Close();
        }
        response = $"{sheetName} has been deleted.";
        return response;
    }

    [McpServerTool(Name = "excel_run_macro"), Description("Run a macro of the specified Excel file.")]
    public static string RunMacro([Description("The full path of the Excel file.")] string fullName
    , [Description("The name of macro.")] string macroName
    , [Description("The parameters of macro. The maximum number is 30.")] string[]? macroParameters = null
    , [Description("Save the file after executing the macro.")] bool save = true
    , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, false, password);
            try
            {
                response = session.RunMacro(macroName, macroParameters);
                wk.Close(save);
            }
            catch (Exception ex)
            {
                throw new McpException(ex.Message);
            }

        }
        if (string.IsNullOrEmpty(response))
        {
            response = $"{macroName} has been called.";
        }
        else
        {
            response = $"The result of {macroName}: {response}";
        }

        return response;
    }

    [McpServerTool(Name = "excel_write"), Description("Write data into a cell or a range of cells of the specified worksheet to an Excel file.")]
    public static string Write([Description("The full path of the Excel file. It will be created if not exist.")] string fullName
    , [Description("The sheet name of the Excel file. It will be created if not exist.")] string sheetName = "Sheet1"
    , [Description("The data that needs to be written in.")] string[][]? data = null
    , [Description("The first column as a letter where the data is written.(such as A)")] string startColumn = "A"
    , [Description("The first row number where the data is written.")] decimal startRow = 1
    , [Description("The password of the Excel file, if there is one.")] string? password = null
    , [Description("Force overwrite to create a new one when the file exists.")] bool forceOverwriteFile = false
    , [Description("Force overwrite to create a new one when the sheet exists.")] bool forceOverwriteSheet = false)
    {
        StringBuilder sb = new StringBuilder();
        using (var session = new ExcelSession())
        {
            fullName = session.CheckFullName(fullName, false, true);
            Excel.Workbook wk;
            Excel.Worksheet sh;
            bool newWk = false;
            bool newSh = false;
            if (File.Exists(fullName) && !forceOverwriteFile)
            {
                wk = session.OpenWorkbook(fullName, false, password);
            }
            else
            {
                newWk = true;
                wk = session.AddWorkbook(sheetName);
            }
            try
            {
                sh = session.GetSheet(wk, sheetName);
            }
            catch
            {
                newSh = true;
                sh = session.AddSheet(wk, sheetName);
            }
            if(!newWk && !newSh && forceOverwriteSheet)
            {
                newSh = true;
                var shNew = session.AddSheet(wk);
                sh.Delete();
                sh = shNew;
                sh.Name = sheetName;
            }
            if(data != null)
            {
                var range = sh.Range[$"{startColumn}{startRow}"];
                session.RegisterComObject(range);
                for (var i = 0; i < data.Length; i++)
                {
                    for (var j = 0; j < data[i].Length; j++)
                    {
                        var cell = range.Offset[i, j];
                        session.RegisterComObject(cell);
                        cell.Value = data[i][j];
                    }
                }
            }
            if (newWk)
            {
                var format = Excel.XlFileFormat.xlOpenXMLWorkbook;
                switch (Path.GetExtension(fullName).ToLowerInvariant())
                {
                    case ".xlsx":
                        format = Excel.XlFileFormat.xlOpenXMLWorkbook;
                        break;
                    case ".xlsm":
                        format = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                        break;
                    case ".xls":
                        format = Excel.XlFileFormat.xlExcel8;
                        break;
                }
                if (string.IsNullOrEmpty(password))
                {
                    wk.SaveAs(fullName, format);
                }
                else
                {
                    wk.SaveAs(fullName, format, password);
                }

            }
            else
            {
                wk.Save();
            }
            if (newWk)
            {
                sb.AppendLine("Successfully saved to a new file.");
            }
            else
            {
                sb.AppendLine("Successfully saved to an existing file.");
            }
            if (newSh)
            {
                sb.AppendLine("Successfully created a new sheet.");
            }
            wk.Close();
        }

        return sb.ToString();
    }
}
