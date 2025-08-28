using OfficeServer.Tools;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;

namespace TestTools;

[TestClass]
public class OutlookTest : TestBase
{

    [TestMethod]
    public void TestGetInboxMailCount()
    {

        var response = OutlookTools.GetInboxMailCount();
        TestContext.WriteLine(response);

    }

    [TestMethod]
    [DataRow(1, 1)]
    [DataRow(2, 1)]
    [DataRow(3, 1)]
    [DataRow(4, 1)]
    [DataRow(5, 1)]
    public void TestReadInboxMails(int startItem, int maxItems)
    {
        var response = OutlookTools.ReadInboxMails(startItem, maxItems);
        TestContext.WriteLine(JsonConvert.SerializeObject(response));
    }

    [TestMethod]
    [DataRow(-2)]
    [DataRow(-1)]
    [DataRow(0)]
    [DataRow(1)]
    [DataRow(2)]
    public void TestReadAppointments(int withinDays)
    {
        var response = OutlookTools.ReadAppointments(withinDays);
        TestContext.WriteLine(JsonConvert.SerializeObject(response));
    }
    [TestMethod]
    [DataRow("", null, 1)]
    [DataRow("", new string[] { "onedrive" }, 1)]
    [DataRow("云环境", null, 1)]
    public void TestFindInboxMails(string searchValue, string[] senders, int maxItems)
    {
        var response = OutlookTools.FindInboxMails(searchValue, senders, maxItems);
        TestContext.WriteLine(JsonConvert.SerializeObject(response));
    }

    [TestMethod]
    [DataRow("test1", new string[] { "user1@example.com" }, "test body", false, null, null, true)]
    [DataRow("test2", new string[] { "user1@example.com" }, "<font color='red'>red</font>", true, null, null, true)]
    [DataRow("test3", new string[] { "user10@example.com", "user12@example.com" }, "test\n body2222", false, new string[] { "user1@example.com", "user2@example.com" }, new string[] { "user3@example.com", "user4@example.com" }, true)]
    public void TestSendMail(string subject, string[] sendTo, string body, bool useHtml, string[] sendCc, string[] sendBcc, bool displayBeforeSend)
    {
        var response = OutlookTools.SendMail(subject, sendTo, body, useHtml, sendCc, sendBcc, displayBeforeSend);
        TestContext.WriteLine(response);
    }

    [TestMethod]
    [DataRow("test appt1", new string[] { "user1@example.com", "user2@example.com" }, 30, "test location", "test appt1 body", true)]
    public void TestSendAppointment(string subject, string[] participants, int duration, string? location, string? body, bool displayBeforeSend)
    {
        var startTime = DateTime.Now.AddDays(1);
        var response = OutlookTools.SendAppointment(subject, participants, startTime, duration, location, body, displayBeforeSend);
        TestContext.WriteLine(response);
    }

}