using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;

namespace TestTools;

[TestClass]
public class PowerPointTest : TestBase
{

    [TestMethod]
    [DataRow("pr1.ppt", null)]
    [DataRow("pr1.pptx", null)]
    [DataRow("pr1.pptm", null)]
    [DataRow("pr1.pptx", "223")]
    public void TestGetSlideCount(string fileName, string password)
    {

        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = PowerPointTools.GetSlideCount(fullName, password);
        TestContext.WriteLine(response);

    }


    [TestMethod]
    [DataRow(new string[] { "pr1.ppt", "pr1.pptx" }, "test", true, true, null)]
    [DataRow(new string[] { "pr1.ppt", "pr1.pptx" }, "test", true, false, null)]
    [DataRow(new string[] { "pr1.ppt", "pr1.pptx" }, "test", false, true, null)]
    [DataRow(new string[] { "pr1.ppt", "pr1.pptx" }, "test", false, false, null)]
    public void TestFind(string[] fileNameList, string searchValue, bool matchPart, bool ignoreCase, string password)
    {

        var fullNames = fileNameList.Select(p => Path.Combine(TestDataDirectory, p)).ToArray();
        var response = PowerPointTools.Find(fullNames, searchValue, matchPart, ignoreCase, password);
        TestContext.WriteLine(response);
    }

    [TestMethod]
    [DataRow("")]
    [DataRow("pr2.ppt")]
    public void TestFileNameCheck(string name)
    {
        var hasError = false;
        var fullName = Path.Combine(TestDataDirectory, name);
        try
        {
            PowerPointTools.GetSlideCount(fullName);
        }
        catch (Exception ex)
        {
            hasError = true;
            TestContext.WriteLine(ex.Message);
        }

        Assert.IsTrue(hasError);


    }

    [TestMethod]
    [DataRow("pr1.ppt", 2, 2, null)]
    [DataRow("pr1.ppt", 1, null, null)]
    [DataRow("pr1.pptx", 1, 1, null)]
    [DataRow("pr1.pptm", 1, null, null)]
    public void TestRead(string fileName, int fromPage, int? toPage, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var content = PowerPointTools.Read(fullName, fromPage, toPage, password);
        TestContext.WriteLine(JsonConvert.SerializeObject(content));
    }


    [TestMethod]
    [DataRow("pr1.pptm", "test1", null, false, null)]
    [DataRow("pr1.pptm", "test2", null, false, null)]
    [DataRow("pr1.pptm", "test3", new string[] { "1", "2", "3" }, false, null)]
    public void TestRunMacro(string fileName, string macroName, string[]? macroParameters, bool save, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        response = PowerPointTools.RunMacro(fullName, macroName, macroParameters, save, password);
        if (macroParameters == null || macroParameters.Length == 0)
        {
            Assert.AreEqual($"{macroName} has been called.", response);
        }
        else
        {
            Assert.AreEqual($"The result of {macroName}: {string.Join("", macroParameters)}", response);
        }

    }


}