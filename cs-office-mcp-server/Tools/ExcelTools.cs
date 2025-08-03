using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using ModelContextProtocol;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class ExcelTools

{
    /// <summary>
    /// Create a new Excel Application instance.
    /// </summary>
    /// <param name="visible">Visible</param>
    /// <param name="displayAlerts">Display warning message</param>
    /// <returns></returns>
    private static Excel.Application CreateApp(bool visible = false, bool displayAlerts = false)
    {
        var app = new Excel.Application();
        app.Visible = visible;
        app.DisplayAlerts = displayAlerts;
        return app;
    }

    /// <summary>
    /// Replace the '/' with '\', and check if the file is valid. Finally, return the replaced path.
    /// </summary>
    /// <param name="fullName">The full path of the Excel file</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    private static string CheckFullName(string fullName)
    {
        fullName = fullName ?? string.Empty;
        fullName = fullName.Replace("/", @"\");
        if (!File.Exists(fullName))
        {
            throw new McpException($"{fullName} not exist.");
        }
        var fileInfo = new FileInfo(fullName);
        var ext = fileInfo.Extension ?? string.Empty;
        ext = ext.ToLower();
        var allowedList = new string[] { "xls", "xlsx", "xlsm" };
        var isExcel = false;
        foreach (var item in allowedList) { 
            if($".{item}" == ext)
            {
                isExcel = true;
                break;
            }
        }
        if (!isExcel) {
            throw new McpException($"{fullName} is not a Excel file.\nCurrently supported formats are [{string.Join(",", allowedList)}].");
        }

        return fullName;
    }
    /// <summary>
    /// Open a Excel file.
    /// </summary>
    /// <param name="app">Excel Application</param>
    /// <param name="fullName">The full path of the Excel file to open</param>
    /// <param name="readOnly">ReadOnly mode</param>
    /// <param name="password">The password</param>
    /// <returns></returns>
    private static Excel.Workbook Open(Excel.Application app, string fullName, bool readOnly, string password)
    {
        Excel.Workbook wk = null;
        if (string.IsNullOrEmpty(password))
        {
            wk = app.Workbooks.Open(fullName, Excel.XlUpdateLinks.xlUpdateLinksNever, readOnly, Type.Missing);
        }
        else
        {
            wk = app.Workbooks.Open(fullName, Excel.XlUpdateLinks.xlUpdateLinksNever, readOnly, Type.Missing, password);
        }
        return wk;
    }

    /// <summary>
    /// Get worksheet by name.
    /// </summary>
    /// <param name="wk">Workbook</param>
    /// <param name="sheetName">Sheet name</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    private static Excel.Worksheet GetSheet(Excel.Workbook wk,string sheetName)
    {
        var sheets = new List<string>();
        foreach (Excel.Worksheet sheet in wk.Sheets)
        {
            sheets.Add(sheet.Name);
        }
        if (!sheets.Contains(sheetName))
        {
            throw new McpException($"{sheetName} not exist in {wk.FullName}.");
        }
        return wk.Sheets[sheetName] as Excel.Worksheet;
    }

    /// <summary>
    /// Get range of worksheet based on the specified start and end columns and rows.
    /// </summary>
    /// <param name="sh">Worksheet</param>
    /// <param name="startColumn">Start column</param>
    /// <param name="startRow">Start row number</param>
    /// <param name="endColumn">End column</param>
    /// <param name="endRow">End row number</param>
    /// <returns></returns>
    private static Excel.Range GetRange(Excel.Worksheet sh, string startColumn, decimal startRow, string endColumn, decimal? endRow)
    {
        var usedRange = sh.UsedRange;
        var startRange = sh.Range[$"{startColumn}{startRow}"];
        var endRange = sh.Range[$"{startColumn}{startRow}"];
        if (string.IsNullOrEmpty(endColumn) && !endRow.HasValue)
        {
            // If both endColumn and endRow are not specified, use xlToRight and xlDown
            endRange = endRange.End[Excel.XlDirection.xlToRight].End[Excel.XlDirection.xlDown];
        }
        else if (!string.IsNullOrEmpty(endColumn) && !endRow.HasValue)
        {
            // If only endColumn is specified, use xlDown
            endRange = sh.Range[$"{endColumn}{startRow}"].End[Excel.XlDirection.xlDown];
        }
        else if (string.IsNullOrEmpty(endColumn) && endRow.HasValue)
        {
            // If only endRow is specified, use xlToRight
            endRange = sh.Range[$"{startColumn}{endRow.Value}"].End[Excel.XlDirection.xlToRight];
        }
        else
        {
            // If both are specified, use the specified values
            endRange = sh.Range[$"{endColumn}{endRow.Value}"];
        }
        if (endRange.Row == sh.Rows.Count) { 
            endRange = sh.Cells[startRange.Row, endRange.Column];
        }
        if (endRange.Column == sh.Columns.Count)
        {
            endRange = sh.Cells[endRange.Row, startRange.Column];
        }
        return sh.Range[startRange,endRange];
    }

    [McpServerTool(Name = "excel_get_sheets"), Description("Get all the sheet names of the specified Excel file.")]
    public static List<string> GetSheets([Description("The full path of the Excel file.")] string fullName, [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        fullName = CheckFullName(fullName);
        Excel.Application app = null;
        var sheets = new List<string>();
        try
        {
            app = CreateApp();
            var wk = Open(app, fullName, true, password);
            foreach (Excel.Worksheet sheet in wk.Sheets) {
                sheets.Add(sheet.Name);
            }
        }
        catch (Exception e)
        {
            throw new McpException(e.Message);
        }
        finally {
            if (app != null)
            {
                app.Quit();
            }
        }

        return sheets;
    }

    [McpServerTool(Name = "excel_read"), Description("Read the value of a cell or a range of cells from the specified worksheet.")]
    public static Dictionary<string, object> Read([Description("The full path of the Excel file.")] string fullName
        , [Description("The sheet name of the Excel file.")] string sheetName
        , [Description("The first column as a letter.(such as A)")] string startColumn = "A"
        , [Description("The first row number.")] decimal startRow = 1
        , [Description("The last column as a letter.(such as Z) If empty, then use xlToRight relative to startColumn")] string? endColumn = null
        , [Description("The last row number. If empty, then use xlDown relative to startRow")] decimal? endRow = null
        , [Description("The password of the Excel file, if there is one.")] string password = null)
    {
        fullName = CheckFullName(fullName);
        Excel.Application app = null;
        var values = new Dictionary<string, object>();
        try
        {
            app = CreateApp();
            var wk = Open(app, fullName, true, password);
            var sh = GetSheet(wk, sheetName);
            var range = GetRange(sh, startColumn, startRow, endColumn, endRow);
            
            foreach (Excel.Range r in range) {
                values[r.Address.Replace("$", "")] = r.Value;
            }
        }
        catch (Exception e)
        {
            throw new McpException(e.Message);
        }
        finally
        {
            if (app != null)
            {
                app.Quit();
            }
        }

        return values;
    }
    [McpServerTool(Name = "excel_read_used_range"), Description("Read the value of used range of cells from the specified worksheet.")]
    public static Dictionary<string, object> ReadUsedRange([Description("The full path of the Excel file.")] string fullName
        , [Description("The sheet name of the Excel file.")] string sheetName
        , [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        fullName = CheckFullName(fullName);
        Excel.Application app = null;
        var values = new Dictionary<string, object>();
        try
        {
            app = CreateApp();
            var wk = Open(app, fullName, true, password);
            var sh = GetSheet(wk, sheetName);
            var range = sh.UsedRange;

            foreach (Excel.Range r in range)
            {
                values[r.Address.Replace("$", "")] = r.Value;
            }
        }
        catch (Exception e)
        {
            throw new McpException(e.Message);
        }
        finally
        {
            if (app != null)
            {
                app.Quit();
            }
        }

        return values;
    }

    [McpServerTool(Name = "excel_write"), Description("Write data into a cell or a range of cells of the specified worksheet.")]
    public static string Write([Description("The full path of the Excel file.")] string fullName
    , [Description("The sheet name of the Excel file.")] string sheetName
    , [Description("The data that needs to be written in.")] string[][] data
    , [Description("The first column as a letter where the data is written.(such as A)")] string startColumn = "A"
    , [Description("The first row number where the data is written.")] decimal startRow = 1
    , [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        if(data == null || data.Length == 0 || data[0].Length == 0)
        {
            throw new McpException("Data to write cannot be empty.");
        }
        fullName = CheckFullName(fullName);
        Excel.Application app = null;
        try
        {
            app = CreateApp();
            var wk = Open(app, fullName, false, password);
            var sh = GetSheet(wk, sheetName);
            var range = sh.Range[$"{startColumn}{startRow}"];

            for(var i = 0; i < data.Length; i++)
            {
                for(var j = 0; j < data[i].Length; j++)
                {
                    var cell = range.Offset[i,j];
                    cell.Value = data[i][j];
                }
            }
            wk.Save();
        }
        catch (Exception e)
        {
            throw new McpException(e.Message);
        }
        finally
        {
            if (app != null)
            {
                app.Quit();
            }
        }

        return "ok";
    }
}
