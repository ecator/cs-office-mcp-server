using ModelContextProtocol;
using ModelContextProtocol.Server;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class WordTools

{

    [McpServerTool(Name = "word_get_page_count"), Description("Get all the number of the pages of the specified Word file.")]
    public static string GetPageCount([Description("The full path of the Word file.")] string fullName
        , [Description("The password of the Word file, if there is one.")] string? password = null)
    {
        var data = new StringBuilder();
        var count = 0;
        data.AppendLine();
        using (var session = new WordSession())
        {
            var doc = session.OpenDocument(fullName, true, password);

            count = session.GetPageCount(doc);
        }
        data.Insert(0, $"Total `{count}` pages in the Word file `{fullName}`.");
        return data.ToString();
    }

    [McpServerTool(Name = "word_read"), Description("Get the text content of the specified Word file.")]
    public static string Read([Description("The full path of the Word file.")] string fullName
        , [Description("The starting page number to read.")] int fromPage = 1
        , [Description("The end page number to read. If it's empty, then read up to the last page.")] int? toPage = null
        , [Description("The password of the Word file, if there is one.")] string? password = null)
    {
        var data = "";
        var pageCount = 0;
        using (var session = new WordSession())
        {
            var doc = session.OpenDocument(fullName, true, password);
            pageCount = session.GetPageCount(doc);
            if (toPage.HasValue && toPage > pageCount)
            {
                throw new McpException($"The end page number {toPage} cannot be greater than the total page count {pageCount}.");
            }
            if (!toPage.HasValue)
            {
                toPage = pageCount;
            }
            var pages = session.GetPageText(doc);
            data = string.Join(Environment.NewLine, pages.GetRange(fromPage - 1, toPage.Value - fromPage + 1));
        }
        return data;
    }

    [McpServerTool(Name = "word_find"), Description("Find value from Word files.")]
    public static string Find([Description("The list of full path of Word files that need to be searched for.")] string[] fullNameList
    , [Description(@"The value to be searched for which can use wildcard characters like ?(any single character), *(any number of characters), \ followed by ?, *, or \(a question mark, asterisk, or backslash).")] string searchValue
    , [Description("Match against any part of part of a larger word when true. Match against the entire words of the search text when false.")] bool matchPart = true
    , [Description("Ignoring lower case and upper case differences when true. Case insensitive when false.")] bool ignoreCase = true
    , [Description("The password of the Word files, if there is one and all are the same.")] string? password = null)
    {
        var data = new StringBuilder();
        var foundData = new StringBuilder();
        var line = new string[3];
        var totalCount = 0;
        var count = 0;
        if (fullNameList == null || fullNameList.Length == 0)
        {
            throw new McpException("The full path list of the Word file cannot be empty or null.");
        }
        data.AppendLine();
        data.AppendLine();
        using (var session = new WordSession())
        {

            foreach (var fullName in fullNameList)
            {
                var doc = session.OpenDocument(fullName, true, password);
                var content = doc.Content;
                session.RegisterComObject(content);
                var docEndPosition = content.End;
                count = 0;
                foundData.Clear();
                foundData.AppendLine();
                var find = content.Find;
                session.RegisterComObject(find);
                find.ClearAllFuzzyOptions();
                find.ClearFormatting();
                find.ClearHitHighlight();
                find.Text = searchValue;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Forward = true;
                find.MatchCase = !ignoreCase;
                find.MatchWholeWord = !matchPart;
                find.MatchWildcards = (searchValue.Contains("?") || searchValue.Contains("*") || searchValue.Contains(@"\\")) ? true : false;
                while (find.Execute())
                {
                    totalCount++;
                    count++;
                    Word.Range parent = find.Parent;
                    session.RegisterComObject(parent);
                    var startPosition = parent.Start - 10;
                    if (startPosition < 0)
                    {
                        startPosition = 0;
                    }
                    var endPosition = parent.End + 10;
                    if (endPosition > docEndPosition)
                    {
                        endPosition = docEndPosition;
                    }
                    var findArea = doc.Range(startPosition, endPosition);
                    session.RegisterComObject(findArea);
                    foundData.AppendLine("<result>");
                    foundData.AppendLine(findArea.Text);
                    foundData.AppendLine("</result>");
                }
                if (count > 0)
                {
                    foundData.Insert(0, $"`{count}` results in `{fullName}`:");
                    data.AppendLine(foundData.ToString());

                }
            }

        }
        data.Insert(0, $"Found a total of `{totalCount}` results for `{searchValue}` in all files.");
        return data.ToString();
    }

    [McpServerTool(Name = "word_replace"), Description("Replace value from Word files.")]
    public static string Replace([Description("The list of full path of Word files that need to be searched for.")] string[] fullNameList
, [Description(@"The value to be searched for which can use wildcard characters like ?(any single character), *(any number of characters), \ followed by ?, *, or \(a question mark, asterisk, or backslash).")] string oldValue
, [Description("The new replacement value.")] string newValue
, [Description("Match against any part of part of a larger word when true. Match against the entire words of the search text when false.")] bool matchPart = true
, [Description("Ignoring lower case and upper case differences when true. Case insensitive when false.")] bool ignoreCase = true
, [Description("Replace all matching items when true. Replace only the first matching item when false.")] bool replaceAll = true
, [Description("The password of the Word files, if there is one and all are the same.")] string? password = null)
    {
        var data = new StringBuilder();
        var line = new string[3];
        if (fullNameList == null || fullNameList.Length == 0)
        {
            throw new McpException("The full path list of the Word file cannot be empty or null.");
        }
        using (var session = new WordSession())
        {

            foreach (var fullName in fullNameList)
            {
                var doc = session.OpenDocument(fullName, false, password);
                var content = doc.Content;
                session.RegisterComObject(content);
                var docEndPosition = content.End;
                var find = content.Find;
                session.RegisterComObject(find);
                find.ClearAllFuzzyOptions();
                find.ClearFormatting();
                find.ClearHitHighlight();
                find.Text = oldValue;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Forward = true;
                find.MatchCase = !ignoreCase;
                find.MatchWholeWord = !matchPart;
                find.MatchWildcards = (oldValue.Contains("?") || oldValue.Contains("*") || oldValue.Contains(@"\\")) ? true : false;
                var replacement = find.Replacement;
                session.RegisterComObject(replacement);
                replacement.ClearFormatting();
                replacement.Text = newValue;
                if (find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, replaceAll ? Word.WdReplace.wdReplaceAll : Word.WdReplace.wdReplaceOne))
                {
                    if (replaceAll)
                    {
                        data.AppendLine($"Replaced all `{oldValue}` with `{newValue}` in `{fullName}`.");
                    }
                    else
                    {
                        data.AppendLine($"Replaced the first found `{oldValue}` with `{newValue}` in `{fullName}`.");
                    }

                    doc.Close(true);
                }
                else
                {
                    data.AppendLine($"Not found `{oldValue}` in `{fullName}`.");
                    doc.Close(false);
                }

            }

        }
        
        return data.ToString();
    }

    [McpServerTool(Name = "word_clear"), Description("Clear the whole content of the specified Word file.")]
    public static string Clear([Description("The full path of the Word file.")] string fullName
    , [Description("The password of the Word file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new WordSession())
        {
            var doc = session.OpenDocument(fullName, false, password);
            Word.Range range = doc.Content;
            session.RegisterComObject(range);
            range.Delete();
            doc.Close(true);
            response = $"The whole content of `{fullName}` cleared.";
        }

        return response;
    }

    [McpServerTool(Name = "word_run_macro"), Description("Run a macro of the specified Word file.")]
    public static string RunMacro([Description("The full path of the Word file.")] string fullName
    , [Description("The name of macro.")] string macroName
    , [Description("The parameters of macro. The maximum number is 30.")] string[]? macroParameters = null
    , [Description("Save the file after executing the macro.")] bool save = true
    , [Description("The password of the Word file, if there is one.")] string password = null)
    {
        var response = "";
        using (var session = new WordSession())
        {
            var doc = session.OpenDocument(fullName, false, password);
            var app = doc.Application;
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
                doc.Close(save);
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

    [McpServerTool(Name = "word_write"), Description("Write data into an Word file.")]
    public static string Write([Description("The full path of the Word file. It will be created if not exist.")] string fullName
    , [Description("The data that needs to be written in.")] string data = ""
    , [Description("Append to the end of the document when true. Append to the beginning of the document when false.")] bool insertAfter = true
    , [Description("Append a newline when writing to an existing file and the newline option is true. When data is appended to the end of the document, a newline character is added before the data. When data is prepended to the beginning of the document, a newline character is added after the data.")] bool insertNewline = true
    , [Description("The password of the Word file, if there is one.")] string? password = null
    , [Description("Force overwrite to create a new one when the file exists.")] bool forceOverwriteFile = false)
    {
        string response = "";
        using (var session = new WordSession())
        {
            fullName = session.CheckFullName(fullName, false, true);
            Word.Document doc;
            bool newDoc = false;
            if (File.Exists(fullName) && !forceOverwriteFile)
            {
                doc = session.OpenDocument(fullName, false, password);
                var range = doc.Content;
                if (insertNewline)
                {
                    if (insertAfter)
                    {
                        range.InsertAfter(Environment.NewLine + data);
                    }
                    else
                    {
                        range.InsertBefore(data + Environment.NewLine);
                    }
                }
                else
                {
                    if (insertAfter)
                    {
                        range.InsertAfter(data);
                    }
                    else
                    {
                        range.InsertBefore(data);
                    }
                }
            }
            else
            {
                newDoc = true;
                doc = session.AddDocument(data);
            }

            if (newDoc)
            {
                var format = Word.WdSaveFormat.wdFormatXMLDocument;
                switch (Path.GetExtension(fullName).ToLowerInvariant())
                {
                    case ".doc":
                        format = Word.WdSaveFormat.wdFormatDocument97;
                        break;
                    case ".docm":
                        format = Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled;
                        break;
                    case ".rtf":
                        format = Word.WdSaveFormat.wdFormatRTF;
                        break;
                }
                if (string.IsNullOrEmpty(password))
                {
                    doc.SaveAs(fullName, format);
                }
                else
                {
                    doc.SaveAs(fullName, format, Type.Missing, password);
                }

            }
            else
            {
                doc.Save();
            }
            if (newDoc)
            {
                response = "Successfully saved to a new file.";
            }
            else
            {
                response = "Successfully saved to an existing file.";
            }
            doc.Close();
        }

        return response;
    }
}
