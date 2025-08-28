using ModelContextProtocol;
using ModelContextProtocol.Server;
using System;
using System.Buffers;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OfficeServer.Tools;


/// <summary>
/// Represents information about an email message, including its subject, sender, recipients, and content.
/// </summary>
/// <remarks>This class provides a structured way to access key details of an email message, such as the sender's
/// name and email address, the recipients (To, CC, BCC), the time the message was received, and the message body. It is
/// typically used to encapsulate email metadata for processing or display purposes.</remarks>
public class MailItemInfo
{
    public string Subject { get; set; }
    public string SenderName { get; set; }
    public string SenderEmailAddress { get; set; }
    public string To { get; set; }
    public string CC { get; set; }
    public string BCC { get; set; }
    public DateTime ReceivedTime { get; set; }
    public string Body { get; set; }
}

/// <summary>
/// Represents the details of an appointment, including its subject, location, time, and description.
/// </summary>
/// <remarks>This class provides a simple structure for storing and managing appointment information. It includes
/// properties for the subject, location, start and end times, and an optional description.</remarks>
public class AppointmentInfo
{
    public string Subject { get; set; }
    public string Location { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Body { get; set; }
}


[McpServerToolType]
public static class OutlookTools

{

    [McpServerTool(Name = "outlook_get_inbox_mail_count"), Description("Get all the number of mail items in the inbox of Outlook.")]
    public static string GetInboxMailCount()
    {
        var data = new StringBuilder();
        var count = 0;
        data.AppendLine();
        using (var session = new OutlookSession())
        {
            var inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var items = inbox.Items;
            session.RegisterComObject(items);
            for (var i = 1; i <= items.Count; i++)
            {
                var item = items[i];
                session.RegisterComObject(item);
                if (item is Outlook.MailItem)
                {
                    count++;
                }
            }
        }
        data.Insert(0, $"Total `{count}` mails in the inbox.");
        return data.ToString();
    }

    [McpServerTool(Name = "outlook_read_inbox_mails"), Description("Read the contents of the mails in the inbox of Outlook.")]
    public static Dictionary<string, List<MailItemInfo>> ReadInboxMails([Description(@"The starting position of the item to be read, which starts from 1.")] int startItem = 1
        , [Description(@"Maximum number of mails to be read.")] int maxItems = 10)
    {
        var data = new Dictionary<string, List<MailItemInfo>>();
        data["inbox"] = new List<MailItemInfo>();
        var count = 0;
        using (var session = new OutlookSession())
        {
            var inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var items = inbox.Items;
            session.RegisterComObject(items);
            for (var i = startItem; i <= items.Count && count < maxItems; i++)
            {
                var item = items[i];
                session.RegisterComObject(item);
                if (item is Outlook.MailItem mailItem)
                {
                    count++;
                    var mailInfo = new MailItemInfo
                    {
                        Subject = mailItem.Subject,
                        SenderName = mailItem.SenderName,
                        SenderEmailAddress = mailItem.SenderEmailAddress,
                        To = mailItem.To,
                        CC = mailItem.CC,
                        BCC = mailItem.BCC,
                        ReceivedTime = mailItem.ReceivedTime,
                        Body = mailItem.Body
                    };
                    data["inbox"].Add(mailInfo);
                }
            }
        }
        return data;
    }

    [McpServerTool(Name = "outlook_find_inbox_mails"), Description("Find the contents of the mails in the inbox of Outlook.")]
    public static Dictionary<string, List<MailItemInfo>> FindInboxMails([Description(@"The value to be searched for will be searched in the subject and body. No filtering if empty.")] string searchValue = ""
        , [Description(@"Email or name of senders need to be specified. No filtering if empty.")] string[] senders = null
        , [Description(@"Maximum number of mails to return.")] int maxItems = 10)
    {
        var data = new Dictionary<string, List<MailItemInfo>>();
        data["found_inbox"] = new List<MailItemInfo>();
        var count = 0;
        using (var session = new OutlookSession())
        {
            var inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            var items = inbox.Items;
            session.RegisterComObject(items);
            var foundItems = items;
            if (!string.IsNullOrEmpty(searchValue))
            {
                searchValue = searchValue.Replace("'", "''");
                foundItems = items.Restrict($"@SQL=(\"http://schemas.microsoft.com/mapi/proptag/0x0037001f\" ci_phrasematch '{searchValue}' OR \"http://schemas.microsoft.com/mapi/proptag/0x1000001e\" ci_phrasematch '{searchValue}')");
            }
            for (var i = 1; i <= foundItems.Count && count < maxItems; i++)
            {
                var item = foundItems[i];
                session.RegisterComObject(item);
                if (item is Outlook.MailItem mailItem)
                {
                    var match = false;
                    if (senders == null || senders.Length == 0)
                    {
                        match = true;
                    }
                    else
                    {
                        foreach (var sender in senders)
                        {
                            if ((!string.IsNullOrEmpty(mailItem.SenderEmailAddress) && mailItem.SenderEmailAddress.Equals(sender, StringComparison.OrdinalIgnoreCase))
                                || (!string.IsNullOrEmpty(mailItem.SenderName) && mailItem.SenderName.Replace(" ", "").Replace("　", "").Equals(sender, StringComparison.OrdinalIgnoreCase)))
                            {
                                match = true;
                                break;
                            }
                        }
                    }
                    if (match)
                    {
                        count++;
                        var mailInfo = new MailItemInfo
                        {
                            Subject = mailItem.Subject,
                            SenderName = mailItem.SenderName,
                            SenderEmailAddress = mailItem.SenderEmailAddress,
                            To = mailItem.To,
                            CC = mailItem.CC,
                            BCC = mailItem.BCC,
                            ReceivedTime = mailItem.ReceivedTime,
                            Body = mailItem.Body
                        };
                        data["found_inbox"].Add(mailInfo);
                    }
                }
            }
        }
        return data;
    }

    [McpServerTool(Name = "outlook_read_appointments"), Description("Read the contents of the appointments in the calendar of Outlook.")]
    public static Dictionary<string, List<AppointmentInfo>> ReadAppointments([Description(@"The range of days for appointments. Read past appointments if a negative number used.")] int withinDays = 30)
    {
        var data = new Dictionary<string, List<AppointmentInfo>>();
        data["appointments"] = new List<AppointmentInfo>();
        var count = 0;
        using (var session = new OutlookSession())
        {
            var calendar = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            var items = calendar.Items;
            session.RegisterComObject(items);
            DateTime startTime;
            DateTime endTime;
            if (withinDays < 0)
            {
                startTime = DateTime.Now.AddDays(withinDays);
                endTime = DateTime.Now;
            }
            else
            {
                startTime = DateTime.Now;
                endTime = startTime.AddDays(withinDays);
            }

            var filter = $"[Start] >= '{startTime.ToString("MM/dd/yyyy HH:mm tt", CultureInfo.InvariantCulture)}' AND [Start] <= '{endTime.ToString("MM/dd/yyyy HH:mm tt", CultureInfo.InvariantCulture)}'";
            var appointments = items.Restrict(filter);
            session.RegisterComObject(appointments);
            for (var i = 1; appointments != null && i <= appointments.Count; i++)
            {
                var appointment = appointments[i];
                session.RegisterComObject(appointment);
                var appointmentInfo = new AppointmentInfo
                {
                    Subject = appointment.Subject,
                    Location = appointment.Location,
                    Start = appointment.Start,
                    End = appointment.End,
                    Body = appointment.Body
                };
                data["appointments"].Add(appointmentInfo);
            }
        }
        return data;
    }

    [McpServerTool(Name = "outlook_send_mail"), Description("Send a mail using Outlook.")]
    public static string SendMail([Description(@"The subject of the mail.")] string subject
        , [Description(@"The primary recipient(s) of the mail.")] string[] sendTo
        , [Description(@"The content of the mail body. Can be plain text or HTML, depending on the useHtml.")] string body
        , [Description(@"A boolean value indicating whether the mail body should be formatted as HTML.")] bool useHtml = false
        , [Description(@"The CC recipient(s) of the mail.")] string[] sendCc = null
        , [Description(@"The BCC recipient(s) of the mail.")] string[] sendBcc = null
        , [Description(@"Display the mail before sending, and send it after manual confirmation, otherwise send it directly.")] bool displayBeforeSend = true)
    {
        var response = "The mail has been created and needs to be sent by the user after confirmation.";
        if (!displayBeforeSend)
        {
            response = "The mail has been sent.";
        }
        using (var session = new OutlookSession())
        {
            var mailItem = session.CreateMailItem(subject, sendTo, sendCc, sendBcc, body, useHtml);
            if (displayBeforeSend)
            {
                mailItem.Display();
            }
            else
            {
                mailItem.Send();
            }
        }

        return response;
    }

    [McpServerTool(Name = "outlook_send_appointment"), Description("Send a appointment using Outlook.")]
    public static string SendAppointment([Description(@"The subject of the appointment.")] string subject
        , [Description(@"An array of email addresses representing the participants of the meeting.")] string[] participants
        , [Description(@"The start time of the appointment.")] DateTime startTime
        , [Description(@"The duration of the appointment in minutes. Must be greater than 0.")] int duration
        , [Description(@"The location of the appointment.")] string? location = null
        , [Description(@"The body content of the appointment.")] string? body = null
        , [Description(@"Display the appointment before sending, and send it after manual confirmation, otherwise send it directly.")] bool displayBeforeSend = true)
    {
        var response = "The appointment has been created and needs to be sent by the user after confirmation.";
        if (!displayBeforeSend)
        {
            response = "The appointment has been sent.";
        }
        using (var session = new OutlookSession())
        {
            var appointmentItem = session.CreateAppointmentItem(subject, participants, startTime, duration, location, body);
            if (displayBeforeSend)
            {
                appointmentItem.Display();
            }
            else
            {
                appointmentItem.Send();
            }
        }

        return response;
    }

}
