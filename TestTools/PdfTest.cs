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

    [TestMethod]
    public void TestMerge()
    {
        var file1 = Path.Combine(TestDataDirectory, "Agent Skills.pdf");
        var file2 = Path.Combine(TestDataDirectory, "MCP.pdf");
        var outputFile = Path.Combine(TestDataDirectory, "MergedOutputTest.pdf");

        if (File.Exists(outputFile))
        {
            File.Delete(outputFile);
        }

        try
        {
            var response = PdfTools.Merge(new[] { file1, file2 }, outputFile);
            Assert.IsTrue(File.Exists(outputFile));
            Assert.IsTrue(response.Contains("Successfully merged"));

            var pageCountMsg = PdfTools.GetPageCount(outputFile);
            Assert.IsTrue(pageCountMsg.Contains("Total `4` pages")); // 2 + 2 = 4
        }
        finally
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }
        }
    }

    [TestMethod]
    public void TestExtract()
    {
        var inputFile = Path.Combine(TestDataDirectory, "Agent Skills.pdf");
        var outputFile = Path.Combine(TestDataDirectory, "ExtractOutputTest.pdf");

        if (File.Exists(outputFile))
        {
            File.Delete(outputFile);
        }

        try
        {
            // Test explicit arguments (pages 1 to 1)
            var response = PdfTools.Extract(inputFile, outputFile, 1, 1);
            Assert.IsTrue(File.Exists(outputFile));
            Assert.IsTrue(response.Contains("Successfully extracted pages 1 to 1"));

            var pageCountMsg = PdfTools.GetPageCount(outputFile);
            Assert.IsTrue(pageCountMsg.Contains("Total `1` pages"));

            File.Delete(outputFile);

            // Test default arguments (should extract all pages, fromPage = 1, toPage = null -> 2)
            response = PdfTools.Extract(inputFile, outputFile);
            Assert.IsTrue(File.Exists(outputFile));
            Assert.IsTrue(response.Contains("Successfully extracted pages 1 to 2"));

            pageCountMsg = PdfTools.GetPageCount(outputFile);
            Assert.IsTrue(pageCountMsg.Contains("Total `2` pages"));
        }
        finally
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }
        }
    }
}

