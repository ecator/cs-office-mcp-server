using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeServer.Tools
{
    /// <summary>
    /// Manages an Excel Application instance and its associated COM objects, ensuring proper release.
    /// Implements IDisposable for use with 'using' statements.
    /// </summary>
    public class ExcelSession : Session
    {
        public Excel.Application Application { get; private set; }

        /// <summary>
        /// Initializes a new Excel session.
        /// </summary>
        /// <param name="visible">Whether the Excel application should be visible.</param>
        /// <param name="displayAlerts">Whether Excel should display alerts (e.g., save prompts).</param>
        public ExcelSession(bool visible = false, bool displayAlerts = false)
        {
            try
            {
                Application = new Excel.Application { Visible = visible, DisplayAlerts = displayAlerts };
                RegisterComObject(Application); // Register the application itself
            }
            catch (Exception ex)
            {
                // If application creation fails, ensure nothing is left hanging
                Dispose(true); // Clean up anything that might have been partially created
                throw new McpException($"Failed to create Excel application: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Open a Excel file.
        /// </summary>
        /// <param name="fullName">The full path of the Excel file to open</param>
        /// <param name="readOnly">ReadOnly mode</param>
        /// <param name="password">The password</param>
        /// <returns></returns>
        public Excel.Workbook OpenWorkbook(string fullName, bool readOnly, string password)
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
            foreach (var item in allowedList)
            {
                if ($".{item}" == ext)
                {
                    isExcel = true;
                    break;
                }
            }
            if (!isExcel)
            {
                throw new McpException($"{fullName} is not a Excel file.\nCurrently supported formats are [{string.Join(",", allowedList)}].");
            }
            Excel.Workbook wk = null;
            try
            {
                
                Excel.Workbooks wks = Application.Workbooks;
                RegisterComObject(wks);
                if (string.IsNullOrEmpty(password))
                {
                    wk = wks.Open(fullName, Excel.XlUpdateLinks.xlUpdateLinksNever, readOnly, Type.Missing);
                }
                else
                {
                    wk = wks.Open(fullName, Excel.XlUpdateLinks.xlUpdateLinksNever, readOnly, Type.Missing, password);
                }
            }
            catch (Exception ex)
            {
                throw new McpException($"Failed to open Excel workbook: {ex.Message}", ex);
            }
            return wk;
        }

        /// <summary>
        /// Get worksheet list.
        /// </summary>
        /// <param name="wk">Workbook</param>
        /// <returns></returns>
        /// <exception cref="McpException"></exception>
        public List<Excel.Worksheet> GetSheets(Excel.Workbook wk)
        {
            Excel.Sheets shs = wk.Sheets;
            RegisterComObject(shs);
            List<Excel.Worksheet> worksheets = new List<Excel.Worksheet>();
            foreach(Excel.Worksheet sh in shs)
            {
                RegisterComObject(sh);
                worksheets.Add(sh);
            }

            return worksheets;
        }

        /// <summary>
        /// Get worksheet by name.
        /// </summary>
        /// <param name="wk">Workbook</param>
        /// <param name="sheetName">Sheet name</param>
        /// <returns></returns>
        /// <exception cref="McpException"></exception>
        public Excel.Worksheet GetSheet(Excel.Workbook wk, string sheetName)
        {
            Excel.Sheets shs = wk.Sheets;
            RegisterComObject(shs);
            Excel.Worksheet sh = null;
            try
            {
                sh = shs[sheetName];
            }
            catch (Exception ex)
            {
                throw new McpException($"{sheetName} not exist in {wk.FullName}.", ex);
            }

            RegisterComObject(sh);
            return sh;
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
        public Excel.Range GetRange(Excel.Worksheet sh, string startColumn, decimal startRow, string endColumn, decimal? endRow)
        {
            var usedRange = sh.UsedRange;
            RegisterComObject(usedRange);
            var startRange = sh.Range[$"{startColumn}{startRow}"];
            RegisterComObject(startRange);
            var endRange = sh.Range[$"{startColumn}{startRow}"];
            RegisterComObject(endRange);
            if (string.IsNullOrEmpty(endColumn) && !endRow.HasValue)
            {
                // If both endColumn and endRow are not specified, use xlToRight and xlDown
                endRange = endRange.End[Excel.XlDirection.xlToRight];
                RegisterComObject(endRange);
                endRange = endRange.End[Excel.XlDirection.xlDown];
                RegisterComObject(endRange);
            }
            else if (!string.IsNullOrEmpty(endColumn) && !endRow.HasValue)
            {
                // If only endColumn is specified, use xlDown
                endRange = sh.Range[$"{endColumn}{startRow}"];
                RegisterComObject(endRange);
                endRange = endRange.End[Excel.XlDirection.xlDown];
                RegisterComObject(endRange);
            }
            else if (string.IsNullOrEmpty(endColumn) && endRow.HasValue)
            {
                // If only endRow is specified, use xlToRight
                endRange = sh.Range[$"{startColumn}{endRow.Value}"];
                RegisterComObject(endRange);
                endRange = endRange.End[Excel.XlDirection.xlToRight];
                RegisterComObject(endRange);
            }
            else
            {
                // If both are specified, use the specified values
                endRange = sh.Range[$"{endColumn}{endRow.Value}"];
                RegisterComObject(endRange);
            }
            var rows = sh.Rows;
            RegisterComObject(rows);
            if (endRange.Row == rows.Count)
            {
                var cells = sh.Cells;
                RegisterComObject(cells);
                endRange = cells[startRange.Row, endRange.Column];
                RegisterComObject(endRange);
            }
            var cols = sh.Columns;
            RegisterComObject(cols);
            if (endRange.Column == cols.Count)
            {
                var cells = sh.Cells;
                RegisterComObject(cells);
                endRange = sh.Cells[endRange.Row, startRange.Column];
                RegisterComObject(endRange);
            }
            var range = sh.Range[startRange, endRange];
            RegisterComObject(range);
            return range;
        }

        /// <summary>
        /// Disposes of the Excel session and releases all COM objects.
        /// </summary>
        /// <param name="disposing">True if called from Dispose(), false if called from finalizer.</param>
        protected override void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // Release managed resources here if any
            }

            // Release unmanaged resources (COM objects)
            if (Application != null)
            {
                try
                {
                    // Before quitting, set DisplayAlerts to false to avoid "Save changes?" prompts
                    Application.DisplayAlerts = false;
                    Application.Quit();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Warning: Failed to quit Excel application: {ex.Message}");
                }
                finally
                {
                    base.Dispose(disposing); // Call base Dispose to release COM objects
                }
            }
        }
    }
}
