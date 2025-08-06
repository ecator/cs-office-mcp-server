using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class ExcelTools

{

    [McpServerTool(Name = "excel_get_sheets"), Description("Get all the sheet names of the specified Excel file.")]
    public static List<string> GetSheets([Description("The full path of the Excel file.")] string fullName, [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        List<string> sheets = new List<string>();
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, true, password);
            
            foreach (Excel.Worksheet sheet in session.GetSheets(wk))
            {
                sheets.Add(sheet.Name);
            }
        }

        return sheets;
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

    [McpServerTool(Name = "excel_write"), Description("Write data into a cell or a range of cells of the specified worksheet.")]
    public static string Write([Description("The full path of the Excel file.")] string fullName
    , [Description("The sheet name of the Excel file.")] string sheetName
    , [Description("The data that needs to be written in.")] string[][] data
    , [Description("The first column as a letter where the data is written.(such as A)")] string startColumn = "A"
    , [Description("The first row number where the data is written.")] decimal startRow = 1
    , [Description("The password of the Excel file, if there is one.")] string? password = null)
    {
        using (var session = new ExcelSession())
        {
            var wk = session.OpenWorkbook(fullName, false, password);
            var sh = session.GetSheet(wk, sheetName);
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
            wk.Save();
        }

        return "ok";
    }
}
