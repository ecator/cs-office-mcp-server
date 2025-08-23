using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;

namespace OfficeServer.Tools;

/// <summary>
/// Manages an PowerPoint Application instance and its associated COM objects, ensuring proper release.
/// Implements IDisposable for use with 'using' statements.
/// </summary>
public class PowerPointSession : Session<PowerPoint.Application>
{

    protected string[] allowedExtList = new string[] { "ppt", "pptx", "pptm", "rtf" };

    /// <summary>
    /// Checks the validity of a full file name for an PowerPoint file.
    /// </summary>
    /// <param name="fullName">The full path of the PowerPoint file.</param>
    /// <param name="needExist">Check if the file exists.</param>
    /// <param name="checkFolder">Check whether the directory of the file exists.</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public string CheckFullName(string fullName, bool needExist = true, bool checkFolder = true)
    {
        if (string.IsNullOrEmpty(fullName))
        {
            throw new McpException($"The file name can not be empty.");
        }
        fullName = fullName.Replace("/", @"\");
        if (!Regex.IsMatch(fullName, @"^(\\|[A-Za-z]:\\).+"))
        {
            throw new McpException($"The file name must be an absolute path or a network file.");
        }
        if (needExist && !File.Exists(fullName))
        {
            throw new McpException($"{fullName} not exist.");
        }
        if (checkFolder && !Directory.Exists(Path.GetDirectoryName(fullName)))
        {
            throw new McpException($"The folder of {fullName} not exist.");
        }

        if (!Regex.IsMatch(fullName, string.Format(@".+\.({0})$", string.Join("|", allowedExtList)), RegexOptions.IgnoreCase))
        {
            throw new McpException($"{fullName} is not a valid PowerPoint file.\nCurrently supported formats are [{string.Join(",", allowedExtList)}].");
        }
        return fullName;
    }
    /// <summary>
    /// Initializes a new PowerPoint session.
    /// </summary>
    /// <param name="visible">Whether the PowerPoint application should be visible.</param>
    /// <param name="displayAlerts">Whether PowerPoint should display alerts (e.g., save prompts).</param>
    public PowerPointSession(bool visible = true, bool displayAlerts = false)
    {
        try
        {
            Application = new PowerPoint.Application();
            Application.Visible = visible ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            Application.DisplayAlerts = displayAlerts ? PowerPoint.PpAlertLevel.ppAlertsAll : PowerPoint.PpAlertLevel.ppAlertsNone;
            RegisterComObject(Application); // Register the application itself
        }
        catch (Exception ex)
        {
            // If application creation fails, ensure nothing is left hanging
            Dispose(true); // Clean up anything that might have been partially created
            throw new McpException($"Failed to create PowerPoint application: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Open a PowerPoint file.
    /// </summary>
    /// <param name="fullName">The full path of the PowerPoint file to open</param>
    /// <param name="readOnly">ReadOnly mode</param>
    /// <param name="password">The password</param>
    /// <returns></returns>
    public PowerPoint.Presentation OpenPresentation(string fullName, bool readOnly, string password)
    {
        fullName = CheckFullName(fullName);
        PowerPoint.Presentation pr = null;
        try
        {

            PowerPoint.Presentations prs = Application.Presentations;
            RegisterComObject(prs);
            if (string.IsNullOrEmpty(password))
            {
                pr = prs.Open(fullName, readOnly ? MsoTriState.msoTrue : MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            }
            else
            {
                pr = prs.Open($"{fullName}::{password}", readOnly ? MsoTriState.msoTrue : MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            }
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to open PowerPoint: {ex.Message}", ex);
        }
        RegisterComObject(pr);
        return pr;
    }

    /// <summary>
    /// Get the number of slides of the presentation.
    /// </summary>
    /// <param name="pr">Presentation</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public int GetSlideCount(PowerPoint.Presentation pr)
    {
        var slides = pr.Slides;
        RegisterComObject(slides);

        return slides.Count;
    }

    /// <summary>
    /// Get the text of shapes.
    /// </summary>
    /// <param name="shapes">Shapes</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public Dictionary<string, string> GetShapesText(PowerPoint.Shapes shapes)
    {
        var dic = new Dictionary<string, string>();
        for (var i = 1; i <= shapes.Count; i++)
        {
            PowerPoint.Shape shape = shapes[i];
            RegisterComObject(shape);
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                var textFrame = shape.TextFrame;
                RegisterComObject(textFrame);
                if (textFrame.HasText == MsoTriState.msoTrue)
                {
                    var textRange = textFrame.TextRange;
                    RegisterComObject(textRange);
                    dic[shape.Name] = textRange.Text;
                }

            }
            if (shape.Type == MsoShapeType.msoTable)
            {
                var table = shape.Table;
                RegisterComObject(table);
                var rows = table.Rows;
                RegisterComObject(rows);
                var columns = table.Columns;
                RegisterComObject(columns);
                var tableText = new StringBuilder();
                var line = new string[columns.Count];
                for (var r = 1; r <= rows.Count; r++)
                {
                    if (r == 2)
                    {
                        Array.Fill(line, "---");
                        tableText.AppendLine(string.Join("|", line));
                    }
                    for (var c = 1; c <= columns.Count; c++)
                    {
                        var cell = table.Cell(r, c);
                        RegisterComObject(cell);
                        var cellShape = cell.Shape;
                        RegisterComObject(cellShape);
                        var cellText = string.Empty;
                        if (cellShape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            var cellTextFrame = cellShape.TextFrame;
                            RegisterComObject(cellTextFrame);
                            if (cellTextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var cellTextRange = cellTextFrame.TextRange;
                                RegisterComObject(cellTextRange);
                                cellText = EscapeMarkdownTableValue(cellTextRange.Text);
                            }
                        }
                        line[c - 1] = cellText;
                    }
                    tableText.AppendLine(string.Join("|", line));
                }
                dic[shape.Name] = tableText.ToString();
            }
        }

        return dic;
    }

    /// <summary>
    /// Get the text of shapes.
    /// </summary>
    /// <param name="shapes">Shapes</param>
    /// <param name="searchValue">The value to be searched for</param>
    /// <param name="matchPart">Match against any part of part of a larger word</param>
    /// <param name="ignoreCase">Ignoring lower case and upper case differences.</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public Dictionary<string, string> FindShapesText(PowerPoint.Shapes shapes, string searchValue, bool matchPart = true, bool ignoreCase = true)
    {
        var dic = new Dictionary<string, string>();
        for (var i = 1; i <= shapes.Count; i++)
        {
            PowerPoint.Shape shape = shapes[i];
            RegisterComObject(shape);
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {
                var textFrame = shape.TextFrame;
                RegisterComObject(textFrame);

                if (textFrame.HasText == MsoTriState.msoTrue)
                {
                    var textRange = textFrame.TextRange;
                    RegisterComObject(textRange);
                    var foundRange = textRange.Find(searchValue, 0, ignoreCase ? MsoTriState.msoFalse : MsoTriState.msoTrue, matchPart ? MsoTriState.msoFalse : MsoTriState.msoTrue);
                    if (foundRange != null)
                    {
                        RegisterComObject(foundRange);
                        dic[shape.Name] = textRange.Text;
                    }

                }

            }
            if (shape.Type == MsoShapeType.msoTable)
            {
                var table = shape.Table;
                RegisterComObject(table);
                var rows = table.Rows;
                RegisterComObject(rows);
                var columns = table.Columns;
                RegisterComObject(columns);
                var tableText = new StringBuilder();
                var line = new string[columns.Count];
                var found = false;
                for (var r = 1; r <= rows.Count; r++)
                {
                    if (r == 2)
                    {
                        Array.Fill(line, "---");
                        tableText.AppendLine(string.Join("|", line));
                    }
                    for (var c = 1; c <= columns.Count; c++)
                    {
                        var cell = table.Cell(r, c);
                        RegisterComObject(cell);
                        var cellShape = cell.Shape;
                        RegisterComObject(cellShape);
                        var cellText = string.Empty;
                        if (cellShape.HasTextFrame == MsoTriState.msoTrue)
                        {
                            var cellTextFrame = cellShape.TextFrame;
                            RegisterComObject(cellTextFrame);
                            if (cellTextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var cellTextRange = cellTextFrame.TextRange;
                                RegisterComObject(cellTextRange);
                                cellText = EscapeMarkdownTableValue(cellTextRange.Text);
                                var foundRange = cellTextRange.Find(searchValue, 0, ignoreCase ? MsoTriState.msoFalse : MsoTriState.msoTrue, matchPart ? MsoTriState.msoFalse : MsoTriState.msoTrue);
                                if (foundRange != null)
                                {
                                    RegisterComObject(foundRange);
                                    found = true;
                                }
                            }
                        }
                        line[c - 1] = cellText;
                    }
                    tableText.AppendLine(string.Join("|", line));
                }
                if (found)
                {
                    dic[shape.Name] = tableText.ToString();
                }

            }
        }

        return dic;
    }

}
