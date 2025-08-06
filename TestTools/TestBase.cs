using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace TestTools;

public abstract class TestBase
{

    public string AssemblyDirectory { get=> Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location); }
    public string TestDataDirectory { get =>Path.GetFullPath(Path.Combine(AssemblyDirectory, @"..\..\..\..\TestData")); }
    public TestContext TestContext { get; set; }
   
}