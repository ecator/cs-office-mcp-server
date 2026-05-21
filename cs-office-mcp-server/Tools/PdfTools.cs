using ModelContextProtocol;
using ModelContextProtocol.Server;
using System;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeServer.Tools;

[McpServerToolType]
public static class PdfTools
{
    [McpServerTool(Name = "pdf_get_page_count"), Description("Get all the number of the pages of the specified PDF file.")]
    public static string GetPageCount([Description("The full path of the PDF file.")] string fullName
        , [Description("The password of the PDF file, if there is one.")] string? password = null)
    {
        var count = 0;
        string resolvedName;
        using (var session = new PdfSession())
        {
            resolvedName = session.CheckFullName(fullName);
            var document = session.OpenDocument(fullName, password);
            count = document.NumberOfPages;
        }
        return $"Total `{count}` pages in the PDF file `{resolvedName}`.";
    }

    [McpServerTool(Name = "pdf_read"), Description("Get the text content of the specified PDF file.")]
    public static string Read([Description("The full path of the PDF file.")] string fullName
        , [Description("The starting page number to read.")] int fromPage = 1
        , [Description("The end page number to read. If it's empty, then read up to the last page.")] int? toPage = null
        , [Description("The password of the PDF file, if there is one.")] string? password = null)
    {
        var data = new StringBuilder();
        using (var session = new PdfSession())
        {
            var document = session.OpenDocument(fullName, password);
            int pageCount = document.NumberOfPages;
            if (toPage.HasValue && toPage > pageCount)
            {
                throw new McpException($"The end page number {toPage} cannot be greater than the total page count {pageCount}.");
            }
            if (!toPage.HasValue)
            {
                toPage = pageCount;
            }
            if (fromPage < 1)
            {
                fromPage = 1;
            }

            for (int i = fromPage; i <= toPage.Value; i++)
            {
                var page = document.GetPage(i);
                data.AppendLine(page.Text);
            }
        }
        return data.ToString();
    }

    [McpServerTool(Name = "pdf_find"), Description("Find value from PDF files.")]
    public static string Find([Description("The list of full path of PDF files that need to be searched for.")] string[] fullNameList
    , [Description(@"The value to be searched for which can use wildcard characters like ?(any single character), *(any number of characters), \ followed by ?, *, or \(a question mark, asterisk, or backslash).")] string searchValue
    , [Description("Match against any part of part of a larger word when true. Match against the entire words of the search text when false.")] bool matchPart = true
    , [Description("Ignoring lower case and upper case differences when true. Case insensitive when false.")] bool ignoreCase = true
    , [Description("The password of the PDF files, if there is one and all are the same.")] string? password = null)
    {
        var data = new StringBuilder();
        var totalCount = 0;

        if (fullNameList == null || fullNameList.Length == 0)
        {
            throw new McpException("The full path list of the PDF file cannot be empty or null.");
        }
        data.AppendLine();
        data.AppendLine();

        // Convert wildcard to regex
        var pattern = Regex.Escape(searchValue)
            .Replace(@"\*", ".*")
            .Replace(@"\?", ".");
            
        if (!matchPart)
        {
            pattern = @"\b" + pattern + @"\b";
        }
        
        RegexOptions options = ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None;
        Regex regex = new Regex(pattern, options);

        foreach (var fullName in fullNameList)
        {
            var count = 0;
            var foundData = new StringBuilder();
            foundData.AppendLine();
            string resolvedName;

            using (var session = new PdfSession())
            {
                resolvedName = session.CheckFullName(fullName);
                var document = session.OpenDocument(fullName, password);
                int pageCount = document.NumberOfPages;
                for (int i = 1; i <= pageCount; i++)
                {
                    var page = document.GetPage(i);
                    var text = page.Text;
                    if (string.IsNullOrEmpty(text))
                        continue;
                    
                    var matches = regex.Matches(text);
                    foreach (Match match in matches)
                    {
                        totalCount++;
                        count++;
                        
                        // Get a snippet of text around the match (e.g., up to 20 chars before and after)
                        int start = Math.Max(0, match.Index - 20);
                        int end = Math.Min(text.Length, match.Index + match.Length + 20);
                        var snippet = text.Substring(start, end - start);
                        
                        foundData.AppendLine("<result>");
                        foundData.AppendLine(snippet);
                        foundData.AppendLine("</result>");
                    }
                }
            }

            if (count > 0)
            {
                foundData.Insert(0, $"`{count}` results in `{resolvedName}`:");
                data.AppendLine(foundData.ToString());
            }
        }
        data.Insert(0, $"Found a total of `{totalCount}` results for `{searchValue}` in all files.");
        return data.ToString();
    }
}
