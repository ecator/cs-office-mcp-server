using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModelContextProtocol.Server;
using Excel = Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using ModelContextProtocol;
using System.Text.RegularExpressions;

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
            var app = wk.Application;
            session.RegisterComObject(app);
            var macroParameterCnt = 0;
            if (macroParameters != null && macroParameters.Length > 0)
            {
                macroParameterCnt = macroParameters.Length;
            }
            try
            {
                switch (macroParameterCnt)
                {
                    case 0:
                        response = app.Run(macroName);
                        break;
                    case 1:
                        response = app.Run(macroName, macroParameters[0]);
                        break;
                    case 2:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1]);
                        break;
                    case 3:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2]);
                        break;
                    case 4:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3]);
                        break;
                    case 5:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4]);
                        break;
                    case 6:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5]);
                        break;
                    case 7:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6]);
                        break;
                    case 8:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7]);
                        break;
                    case 9:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8]);
                        break;
                    case 10:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9]);
                        break;
                    case 11:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10]);
                        break;
                    case 12:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11]);
                        break;
                    case 13:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12]);
                        break;
                    case 14:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13]);
                        break;
                    case 15:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14]);
                        break;
                    case 16:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15]);
                        break;
                    case 17:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16]);
                        break;
                    case 18:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17]);
                        break;
                    case 19:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18]);
                        break;
                    case 20:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19]);
                        break;
                    case 21:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20]);
                        break;
                    case 22:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21]);
                        break;
                    case 23:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22]);
                        break;
                    case 24:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23]);
                        break;
                    case 25:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24]);
                        break;
                    case 26:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25]);
                        break;
                    case 27:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26]);
                        break;
                    case 28:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27]);
                        break;
                    case 29:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27], macroParameters[28]);
                        break;
                    case 30:
                        response = app.Run(macroName, macroParameters[0], macroParameters[1], macroParameters[2], macroParameters[3], macroParameters[4], macroParameters[5], macroParameters[6], macroParameters[7], macroParameters[8], macroParameters[9], macroParameters[10], macroParameters[11], macroParameters[12], macroParameters[13], macroParameters[14], macroParameters[15], macroParameters[16], macroParameters[17], macroParameters[18], macroParameters[19], macroParameters[20], macroParameters[21], macroParameters[22], macroParameters[23], macroParameters[24], macroParameters[25], macroParameters[26], macroParameters[27], macroParameters[28], macroParameters[29]);
                        break;
                    default:
                        throw new McpException("Exceeds the maximum number of macro parameters, which is 30.");
                }
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
