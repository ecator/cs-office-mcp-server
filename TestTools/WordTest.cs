using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Text;

namespace TestTools;

[TestClass]
public class WordTest : TestBase
{

    [TestMethod]
    [DataRow("doc1.doc", null)]
    [DataRow("doc1.docx", null)]
    [DataRow("doc1.docm", null)]
    [DataRow("doc1.rtf", null)]
    [DataRow("doc2-password-123.docx", "123")]
    public void TestGetPageCount(string fileName, string password)
    {

        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = WordTools.GetPageCount(fullName, password);
        TestContext.WriteLine(response);

    }


    [TestMethod]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, "动物园", true, true, null)]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, "动物??运营", false, true, null)]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, "动物*", false, true, null)]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, @"\?", true, true, null)]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, "tolist", true, true, null)]
    [DataRow(new string[] { "doc1.doc", "doc1.docx" }, "tolist", true, false, null)]
    public void TestFind(string[] fileNameList, string searchValue, bool matchPart, bool ignoreCase, string password)
    {

        var fullNames = fileNameList.Select(p => Path.Combine(TestDataDirectory, p)).ToArray();
        var response = WordTools.Find(fullNames, searchValue, matchPart, ignoreCase, password);
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
            WordTools.GetPageCount(fullName);
        }
        catch (Exception ex)
        {
            hasError = true;
            TestContext.WriteLine(ex.Message);
        }

        Assert.IsTrue(hasError);


    }

    [TestMethod]
    [DataRow("doc1.doc", 2, 2, null)]
    [DataRow("doc1.doc", 1, null, null)]
    [DataRow("doc1.doc", 2, null, null)]
    public void TestRead(string fileName, int fromPage, int? toPage, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var content = WordTools.Read(fullName, fromPage, toPage, password);
        TestContext.WriteLine(content);
    }

    [TestMethod]
    [DataRow("write-clear-0.doc", null)]
    public void TestClear(string fileName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        var data = DateTime.Now.ToString();
        WordTools.Write(fullName, data, true, true, password, true);
        response = WordTools.Clear(fullName, password);
        Assert.AreEqual($"The whole content of `{fullName}` cleared.", response);
    }

    [TestMethod]
    [DataRow("doc1.docm", "test1", null, false, null)]
    [DataRow("doc1.docm", "test2", new string[] { "1" }, false, null)]
    [DataRow("doc1.docm", "test3", new string[] { "1", "2" }, false, null)]
    [DataRow("doc1.docm", "test4", new string[] { "1", "2", "3", "@@@@@" }, false, null)]
    public void TestRunMacro(string fileName, string macroName, string[]? macroParameters, bool save, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        response = WordTools.RunMacro(fullName, macroName, macroParameters, save, password);
        if (macroParameters == null || macroParameters.Length == 0 || macroName == "test2")
        {
            Assert.AreEqual($"{macroName} has been called.", response);
        }
        else
        {
            Assert.AreEqual($"The result of {macroName}: {string.Join("", macroParameters)}", response);
        }

    }

    [TestMethod]
    [DataRow("write-0.doc", "test1", true, false, null, true)]
    [DataRow("write-0-pass-223.docx", "test2", true, true, "223", false)]
    [DataRow("write-0.docm", "测试2222", false, true, null, false)]
    [DataRow("write-0.rtf", "换行1\n换行1\naaaaaaa", false, false, null, true)]
    public void TestWrite(string fileName, string data, bool insertAfter, bool insertNewline, string password, bool forceOverwriteFile)
    {
        const string RESULT_NEW_FILE = "Successfully saved to a new file.";
        const string RESULT_EXISTING_FILE = "Successfully saved to an existing file.";
        var expectedResponse = string.Empty;
        var response = string.Empty;
        var fullName = Path.Combine(TestDataDirectory, fileName);
        if (File.Exists(fullName))
        {
            File.Delete(fullName);
        }
        response = WordTools.Write(fullName, data, insertAfter, insertNewline, password, forceOverwriteFile);
        expectedResponse = RESULT_NEW_FILE;
        Assert.AreEqual(expectedResponse, response);
        var data2 = DateTime.Now.ToString();
        response = WordTools.Write(fullName, data2, insertAfter, insertNewline, password, forceOverwriteFile);
        if (forceOverwriteFile)
        {
            Assert.AreEqual(expectedResponse, response);
            response = WordTools.Read(fullName, 1, null, password);
            expectedResponse = data2;
            Assert.AreEqual(expectedResponse, response.Trim());
        }
        else
        {
            expectedResponse = RESULT_EXISTING_FILE;
            Assert.AreEqual(expectedResponse, response);
            response = WordTools.Read(fullName, 1, null, password);
            if (insertNewline)
            {
                if (insertAfter)
                {
                    expectedResponse = data + "\r" + data2;
                }
                else
                {
                    expectedResponse = data2 + "\r" + data;
                }
            }
            else
            {
                if (insertAfter)
                {
                    expectedResponse = data + data2;
                }
                else
                {
                    expectedResponse = data2 + data;
                }
            }
            Assert.AreEqual(expectedResponse, response.Trim());
        }
    }

    [TestMethod]
    [DataRow(new string[] { "write-1.docx" }, "动物园", "植物园", true, true, true, null)]
    [DataRow(new string[] { "write-2.docx" }, "动物??运营", "委员会", false, true, false, null)]
    [DataRow(new string[] { "write-3.docx" }, "动物*","海洋馆" ,false, true, false, null)]
    [DataRow(new string[] { "write-4.docx" }, @"\?", "问号", true, true, true, null)]
    [DataRow(new string[] { "write-5.docx" }, "tolist", "TOLIST", true, true, true, null)]
    [DataRow(new string[] { "write-6.docx" }, "tolist", "TOLIST", true, false, true, null)]
    public void TestReplace(string[] fileNameList, string oldValue, string newValue, bool matchPart, bool ignoreCase, bool replaceAll, string password)
    {

        var fullNames = fileNameList.Select(p => Path.Combine(TestDataDirectory, p)).ToArray();
        var data = "动物园是一个很好的地方。动物园有很多动物。动物园的运营很好，但是动物园的运营是外包的。ToList方法可以得到一个列表。?";
        foreach (var fullName in fullNames)
        {
            if (File.Exists(fullName))
            {
                File.Delete(fullName);
            }
            WordTools.Write(fullName, data, true, true, password, true);
            var response = WordTools.Replace(fullNames, oldValue, newValue, matchPart, ignoreCase, replaceAll, password);
            TestContext.WriteLine(response);
            response = WordTools.Read(fullName, 1, null, password);
            TestContext.WriteLine(response);
        }
        
        

       
    }

}