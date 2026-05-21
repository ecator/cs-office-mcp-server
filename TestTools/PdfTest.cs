using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Text;
using System.IO;
using System.Linq;
using System;

namespace TestTools;

[TestClass]
public class PdfTest : TestBase
{

    [TestMethod]
    [DataRow("Agent Skills.pdf", null, 2)]
    [DataRow("MCP.pdf", null, 2)]
    [DataRow("Secret-123456.pdf", "123456", 1)]
    public void TestGetPageCount(string fileName, string password, int expectedPageCount)
    {

        var fullName = Path.Combine(TestDataDirectory, fileName);
        var expected = $"Total `{expectedPageCount}` pages in the PDF file `{fullName}`.";
        var response = PdfTools.GetPageCount(fullName, password);
        Assert.AreEqual(expected, response);

    }


    [TestMethod]
    [DataRow(new string[] { "Agent Skills.pdf" }, "agent", true, true, null)]
    [DataRow(new string[] { "MCP.pdf" }, "Model*", false, true, null)]
    [DataRow(new string[] { "Agent Skills.pdf" }, "tool", true, true, null)]
    [DataRow(new string[] { "Secret-123456.pdf" }, "Secret", true, true, "123456")]
    public void TestFind(string[] fileNameList, string searchValue, bool matchPart, bool ignoreCase, string password)
    {

        var fullNames = fileNameList.Select(p => Path.Combine(TestDataDirectory, p)).ToArray();
        var response = PdfTools.Find(fullNames, searchValue, matchPart, ignoreCase, password);
        TestContext.WriteLine(response);
    }

    [TestMethod]
    [DataRow("")]
    [DataRow("test.txt")]
    public void TestFileNameCheck(string name)
    {
        var hasError = false;
        var fullName = Path.Combine(TestDataDirectory, name);
        try
        {
            PdfTools.GetPageCount(fullName);
        }
        catch (Exception ex)
        {
            hasError = true;
            TestContext.WriteLine(ex.Message);
        }

        Assert.IsTrue(hasError);


    }

    [TestMethod]
    [DataRow("Agent Skills.pdf", 1, 1, null)]
    [DataRow("MCP.pdf", 1, null, null)]
    [DataRow("Secret-123456.pdf", 1, null, "123456")]
    public void TestRead(string fileName, int fromPage, int? toPage, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var content = PdfTools.Read(fullName, fromPage, toPage, password);
        TestContext.WriteLine(content);
    }

}
