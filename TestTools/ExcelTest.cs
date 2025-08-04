using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace TestTools
{
    [TestClass]
    public class ExcelTest
    {

        public string AssemblyDirectory { get=> Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
        public string TestDataDirectory { get =>Path.GetFullPath(Path.Combine(AssemblyDirectory, @"..\..\..\..\TestData")); }
        public TestContext TestContext { get; set; }
        [TestMethod]
        public void TestGetSheets()
        {

            var fullName = Path.Combine(TestDataDirectory, "wk2-password-223.xlsx");
            var sheets = ExcelTools.GetSheets(fullName, "223");
            TestContext.WriteLine(string.Join(",", sheets));

        }

        [TestMethod]
        public void TestRead()
        {
            var fullName = Path.Combine(TestDataDirectory, "wk1.xlsm");
            var password = "";
            var sheetName = "Sheet2";
            var values = ExcelTools.Read(fullName, sheetName, "A", 1, "A", null, password);
            TestContext.WriteLine(string.Join("\n",values.Select(item=>$"{item.Key}={Convert.ToString(item.Value)}")));
            values = ExcelTools.Read(fullName, sheetName, "A", 1, null, null, password);
            TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));
            values = ExcelTools.Read(fullName, sheetName, "A", 1, null, 1, password);
            TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));
            values = ExcelTools.Read(fullName, sheetName, "AZ", 100, null, null, password);
            TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));


        }

        [TestMethod]
        public void TestReadUsedRange()
        {
            var fullName = Path.Combine(TestDataDirectory, "wk0.xls");
            var password = "";
            var sheetName = "Sheet1";
            var values = ExcelTools.ReadUsedRange(fullName, sheetName);
            TestContext.WriteLine(string.Join("\n", values.Select(item => $"{item.Key}={Convert.ToString(item.Value)}")));


        }

        [TestMethod]
        public void TestWrite()
        {
            var fullName = Path.Combine(TestDataDirectory, "wk1.xlsm");
            var password = "";
            var sheetName = "Sheet2";
            var data = new string[][]
            {
                new []{"1","低挫",DateTime.Now.ToString() },
                new []{"2","こんにちは", DateTime.Now.ToString() }
            };
            var ok = ExcelTools.Write(fullName, sheetName, data, "A", 1, password);
            Assert.AreEqual(ok, "ok");


        }
    }
}