using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace OfficeServer.Tools;

/// <summary>
/// Manages an Wor Application instance and its associated COM objects, ensuring proper release.
/// Implements IDisposable for use with 'using' statements.
/// </summary>
public class WordSession : Session<Word.Application>
{

    protected string[] allowedExtList = new string[] { "doc", "docm", "docx", "rtf" };

    /// <summary>
    /// Checks the validity of a full file name for an Word file.
    /// </summary>
    /// <param name="fullName">The full path of the Word file.</param>
    /// <param name="needExist">Check if the file exists.</param>
    /// <param name="checkFolder">Check whether the directory of the file exists.</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public string CheckFullName(string fullName, bool needExist = true, bool checkFolder =true) {
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
            throw new McpException($"{fullName} is not a valid Word file.\nCurrently supported formats are [{string.Join(",", allowedExtList)}].");
        }
        return fullName;
    }
    /// <summary>
    /// Initializes a new Word session.
    /// </summary>
    /// <param name="visible">Whether the Word application should be visible.</param>
    /// <param name="displayAlerts">Whether Word should display alerts (e.g., save prompts).</param>
    public WordSession(bool visible = false, bool displayAlerts = false)
    {
        try
        {
            Application = new Word.Application { Visible = visible, DisplayAlerts = displayAlerts ? Word.WdAlertLevel.wdAlertsAll : Word.WdAlertLevel.wdAlertsNone };
            RegisterComObject(Application); // Register the application itself
        }
        catch (Exception ex)
        {
            // If application creation fails, ensure nothing is left hanging
            Dispose(true); // Clean up anything that might have been partially created
            throw new McpException($"Failed to create Word application: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Open a Word file.
    /// </summary>
    /// <param name="fullName">The full path of the Word file to open</param>
    /// <param name="readOnly">ReadOnly mode</param>
    /// <param name="password">The password</param>
    /// <returns></returns>
    public Word.Document OpenDocument(string fullName, bool readOnly, string password)
    {
        fullName = CheckFullName(fullName);
        Word.Document doc = null;
        try
        {
            
            Word.Documents docs = Application.Documents;
            RegisterComObject(docs);
            if (string.IsNullOrEmpty(password))
            {
                doc = docs.Open(fullName, Type.Missing, readOnly, Type.Missing);
            }
            else
            {
                doc = docs.Open(fullName, Type.Missing, readOnly, Type.Missing, password);
            }
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to open Word workbook: {ex.Message}", ex);
        }
        RegisterComObject(doc);
        return doc;
    }

    /// <summary>
    /// Add a new document.
    /// </summary>
    /// <param name="text">The text content to added initially.</param>
    /// <returns></returns>
    public Word.Document AddDocument(string text)
    {
        Word.Document doc = null;
        try
        {

            Word.Documents docs = Application.Documents;
            RegisterComObject(docs);
            doc = docs.Add();
            Word.Range range = doc.Content;
            RegisterComObject(range);
            range.Text = text;
            RegisterComObject(range);
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to add document: {ex.Message}", ex);
        }
        RegisterComObject(doc);
        return doc;
    }

    /// <summary>
    /// Get the number of pages of the document.
    /// </summary>
    /// <param name="doc">Document</param>
    /// <returns></returns>
    /// <exception cref="McpException"></exception>
    public int GetPageCount(Word.Document doc)
    {
        var pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

        return pageCount;
    }

    /// <summary>
    /// Get text content of pages.
    /// </summary>
    /// <param name="doc">Document</param>
    /// <returns></returns>
    public List<string> GetPageText(Word.Document doc)
    {
        List<string> pageTexts = new List<string>();
        try
        {
            Word.Range content = doc.Content;
            RegisterComObject(content);
            int pageCount = GetPageCount(doc);
            for (int i = 1; i <= pageCount; i++)
            {
                Word.Range pageRange = doc.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, i);
                RegisterComObject(pageRange);
                var pageRangeEnd = content.End;
                if (i < pageCount)
                {
                    var nextPage = doc.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToAbsolute, i + 1);
                    RegisterComObject(nextPage);
                    pageRangeEnd = nextPage.Start;
                }
                pageRange.End = pageRangeEnd;
                string pageText = pageRange.Text;
                pageTexts.Add(pageText);
            }
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to get page text: {ex.Message}", ex);
        }
        return pageTexts;
    }

}
