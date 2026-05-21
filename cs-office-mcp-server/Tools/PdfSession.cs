using ModelContextProtocol;
using System;
using System.IO;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace OfficeServer.Tools;

public class PdfSession : IDisposable
{
    private PdfDocument? _document;

    public string CheckFullName(string fullName, bool needExist = true)
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
        if (!fullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            throw new McpException($"{fullName} is not a valid PDF file.");
        }
        return fullName;
    }

    public PdfDocument OpenDocument(string fullName, string? password = null)
    {
        fullName = CheckFullName(fullName);
        _document = PdfDocument.Open(fullName, new ParsingOptions { Password = password });
        return _document;
    }

    public void Dispose()
    {
        _document?.Dispose();
    }
}
