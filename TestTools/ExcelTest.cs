using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

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
    [DataRow("wk1.xlsm", "Sheet1", null)]
    public void TestReadUsedRange(string fileName, string sheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var values = ExcelTools.ReadUsedRange(fullName, sheetName, password);
        TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));
    }

    [TestMethod]
    [DataRow("write-0.xlsm", "Sheet111", null)]
    [DataRow("write-0.xls", "Sheet123", "123")]
    [DataRow("write-0.xlsx", "Sheet444", "4444")]

    public void TestWrite(string fileName, string sheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var data = new string[][]
        {
            new []{"1","你好",DateTime.Now.ToString() },
            new []{"2","こんにちは", DateTime.Now.ToString() },
            new []{"3","天気がいいから\n散歩ししましょう。", DateTime.Now.ToString() }
        };
        ExcelTools.Save(fullName, sheetName, null, "B", 3, password);
        var ok = ExcelTools.Write(fullName, sheetName, data, "B", 3, password);
        Assert.AreEqual(ok, "ok");

        var values = ExcelTools.Read(fullName, sheetName, "B", 3, null, null, password);
        for (var i = 0; i < data.Length; i++)
        {
            for (var j = 0; j < data[i].Length; j++)
            {
                Assert.AreEqual(Convert.ToString(data[i][j]), Convert.ToString(values[$"{(char)('B' + j)}{3 + i}"]));
            }
        }

    }

    [TestMethod]
    [DataRow("save-1.xlsm", "Sheet1", null)]
    [DataRow("save-1.xls", "Sheet223", "334")]
    [DataRow("save-1.xlsx", "Sheet554", "554")]

    public void TestSave(string fileName, string sheetName, string password)
    {
        var fullName = Path.Combine(TestDataDirectory, fileName);
        var data = new string[][]
        {
            new []{"1","你好",DateTime.Now.ToString() },
            new []{"2","こんにちは", DateTime.Now.ToString() },
            new []{"3", "油断大敵\n危ない。", DateTime.Now.ToString() }
        };
        var ok = ExcelTools.Save(fullName, sheetName, data, "A", 3, password);
        Assert.AreEqual(ok, "ok");

        var values = ExcelTools.Read(fullName, sheetName, "A", 3, null, null, password);
        for (var i = 0; i < data.Length; i++)
        {
            for (var j = 0; j < data[i].Length; j++)
            {
                Assert.AreEqual(Convert.ToString(data[i][j]), Convert.ToString(values[$"{(char)('A' + j)}{3 + i}"]));
            }
        }

    }
}