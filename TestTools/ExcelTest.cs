using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Text;

namespace TestTools;

[TestClass]
public class ExcelTest: TestBase
{

    [TestMethod]
    [DataRow("wk2-password-223.xlsx","223")]
    public void TestGetSheets(string fileName,string password)
    {

        var fullName = Path.Combine(TestDataDirectory, fileName);
        var sheets = ExcelTools.GetSheets(fullName, password);
        TestContext.WriteLine(string.Join(",", sheets));

    }

    [TestMethod]
    [DataRow("wk1.xlsm", null)]
    public void TestGetTables(string fileName, string password)
    {

        var fullName = Path.Combine(TestDataDirectory, fileName);
        var tables = ExcelTools.GetTables(fullName, password);
        foreach (var table in tables)
        {
            var response = ExcelTools.GetTableContent(fullName, table, password);
            TestContext.WriteLine(response);
        }


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
            var sheets = ExcelTools.GetSheets(fullName);
        }catch(Exception ex)
        {
            hasError = true;
            TestContext.WriteLine(ex.Message);
        }

        Assert.IsTrue(hasError);
        

    }

    [TestMethod]
    [DataRow("wk1.xlsm", "Sheet1", "B", 1, "B", 1, null)]
    [DataRow("wk1.xlsm", "Sheet1", "B", 1, null, null, null)]
    [DataRow("wk1.xlsm", "Sheet1", "B", 1, "B", null, null)]
    [DataRow("wk1.xlsm", "Sheet1", "B", 1, null, 2, null)]
    [DataRow("wk1.xlsm", "Sheet1", "C", 1, null, null, null)]
    public void TestRead(string fileName, string sheetName, string startColumn, int startRow, string endColumn, int? endRow, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var values = ExcelTools.Read(fullName, sheetName, startColumn, startRow, endColumn, endRow, password);
        TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));
    }

    [TestMethod]
    [DataRow("write-clear-0.xlsm", "Sheet1", "C", 1, "B", 1, null)]
    [DataRow("write-clear-0.xlsm", "Sheet2", null, null, null, null, null)]
    public void TestClear(string fileName, string sheetName, string startColumn, int? startRow, string endColumn, int? endRow, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        var data = new string[][]
        {
            new []{"1","你好",DateTime.Now.ToString() },
            new []{"2","こんにちは", DateTime.Now.ToString() },
            new []{"3","天気がいいから\n散歩ししましょう。", DateTime.Now.ToString() }
        };
        ExcelTools.Write(fullName, sheetName, data, "A", 1, password);
        response = ExcelTools.Clear(fullName, sheetName, startColumn, startRow, endColumn, endRow, password);
        if (string.IsNullOrEmpty(startColumn))
        {
            Assert.AreEqual("All cells cleared.", response);
        }
        else
        {
            Assert.AreEqual("$B$1:$C$1 cleared.", response);
        }
    }
    [TestMethod]
    [DataRow("write-rename-0.xlsm", "Sheet1", "Sheet1R", null)]
    public void TestRenameSheet(string fileName, string oldSheetName, string newSheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        ExcelTools.Write(fullName, oldSheetName, null, "A", 1, password, true);
        response = ExcelTools.RenameSheet(fullName, oldSheetName, newSheetName, password);
        Assert.AreEqual($"{oldSheetName} has been changed to {newSheetName}.", response);
    }
    [TestMethod]
    [DataRow("write-delete-0.xlsm", "Sheet1", null)]
    public void TestDeleteSheet(string fileName, string sheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        ExcelTools.Write(fullName, sheetName, null, "A", 1, password, true);
        ExcelTools.Write(fullName, sheetName + "1", null, "A", 1, password, false);
        response = ExcelTools.DeleteSheet(fullName, sheetName, password);
        Assert.AreEqual($"{sheetName} has been deleted.", response);
    }

    [TestMethod]
    [DataRow("wk0-macro.xlsm", "test1", null, false, null)]
    [DataRow("wk0-macro.xlsm", "test2", new string[] { "1" }, false, null)]
    [DataRow("wk0-macro.xlsm", "test3", new string[] { "1", "2" }, false, null)]
    [DataRow("wk0-macro.xlsm", "test4", new string[] { "1", "2", "3", "4", "5", "6", "7", "AAAA" }, false, null)]
    public void TestRunMacro(string fileName, string macroName, string[]? macroParameters, bool save, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var response = "";
        response = ExcelTools.RunMacro(fullName, macroName, macroParameters, save, password);
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
    [DataRow("wk1.xlsm", "Sheet1", null)]
    public void TestReadUsedRange(string fileName, string sheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var values = ExcelTools.ReadUsedRange(fullName, sheetName, password);
        TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));
    }

    [TestMethod]
    [DataRow("write-0.xlsm", "Sheet111", null,false,false)]
    [DataRow("write-0.xls", "Sheet123", "123", true, false)]
    [DataRow("write-0.xlsx", "Sheet444", "4444", false, true)]
    [DataRow("write-0.xlsx", "Sheet555", null, true, true)]
    public void TestWrite(string fileName, string sheetName, string password, bool forceOverwriteFile ,bool forceOverwriteSheet)
    {
        const string RESULT_NEW_FILE = "Successfully saved to a new file.";
        const string RESULT_EXISTING_FILE = "Successfully saved to an existing file.";
        const string RESULT_NEW_SHEET = "Successfully created a new sheet.";
        var expectedResponse = new StringBuilder();
        var response = string.Empty;
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var data = new string[][]
        {
            new []{"1","你好",DateTime.Now.ToString() },
            new []{"2","こんにちは", DateTime.Now.ToString() },
            new []{"3","天気がいいから\n散歩ししましょう。", DateTime.Now.ToString() }
        };
        if (File.Exists(fullName))
        {
            File.Delete(fullName);
        }
        response = ExcelTools.Write(fullName, sheetName, null, "B", 3, password, forceOverwriteFile, forceOverwriteSheet);
        expectedResponse.AppendLine(RESULT_NEW_FILE);
        Assert.AreEqual(expectedResponse.ToString(), response);
        response = ExcelTools.Write(fullName, sheetName, data, "B", 3, password, forceOverwriteFile, forceOverwriteSheet);
        expectedResponse.Clear();
        if (forceOverwriteFile)
        {
            expectedResponse.AppendLine(RESULT_NEW_FILE);
        }
        else
        {
            expectedResponse.AppendLine(RESULT_EXISTING_FILE);
            if (forceOverwriteSheet)
            {
                expectedResponse.AppendLine(RESULT_NEW_SHEET);
            }
        }
        Assert.AreEqual(expectedResponse.ToString(), response);

        sheetName = sheetName + "1";
        response = ExcelTools.Write(fullName, sheetName, data, "B", 3, password, forceOverwriteFile, forceOverwriteSheet);
        expectedResponse.Clear();
        if (forceOverwriteFile)
        {
            expectedResponse.AppendLine(RESULT_NEW_FILE);
        }
        else
        {
            expectedResponse.AppendLine(RESULT_EXISTING_FILE);
            expectedResponse.AppendLine(RESULT_NEW_SHEET);
        }
        Assert.AreEqual(expectedResponse.ToString(), response);

        var values = ExcelTools.Read(fullName, sheetName, "B", 3, null, null, password);
        for (var i = 0; i < data.Length; i++)
        {
            for (var j = 0; j < data[i].Length; j++)
            {
                Assert.AreEqual(Convert.ToString(data[i][j]), Convert.ToString(values[$"{(char)('B' + j)}{3 + i}"]));
            }
        }

    }

}