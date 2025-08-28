using ModelContextProtocol;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;

namespace OfficeServer.Tools;

/// <summary>
/// Manages an Outlook Application instance and its associated COM objects, ensuring proper release.
/// Implements IDisposable for use with 'using' statements.
/// </summary>
public class OutlookSession : Session<Outlook.Application>
{

    /// <summary>
    /// Initializes a new Outlook session.
    /// </summary>
    public OutlookSession()
    {
        try
        {
            Application = new Outlook.Application();
            RegisterComObject(Application); // Register the application itself
        }
        catch (Exception ex)
        {
            // If application creation fails, ensure nothing is left hanging
            Dispose(true); // Clean up anything that might have been partially created
            throw new McpException($"Failed to create Outlook application: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Open a Outlook file.
    /// </summary>
    /// <param name="folderType">The type of default folder to return</param>
    /// <returns></returns>
    public Outlook.MAPIFolder GetDefaultFolder(Outlook.OlDefaultFolders folderType)
    {
        var ns = Application.GetNamespace("MAPI");
        Outlook.MAPIFolder folder;
        RegisterComObject(ns);
        try
        {
            folder = ns.GetDefaultFolder(folderType);
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to get folder: {ex.Message}", ex);
        }
        RegisterComObject(folder);
        return folder;
    }

    /// <summary>
    /// Creates a new Outlook mail item with the specified subject, recipients, and body content.
    /// </summary>
    /// <remarks>This method initializes a new Outlook mail item and sets its subject, recipients, and body
    /// content. The <paramref name="sendCc"/> and <paramref name="sendBcc"/> parameters are optional and can be omitted
    /// if not needed. The created mail item is registered as a COM object to ensure proper resource
    /// management.</remarks>
    /// <param name="subject">The subject of the email. Cannot be null or empty.</param>
    /// <param name="sendTo">The primary recipient(s) of the email, specified as an array of string. Cannot be null or empty.</param>
    /// <param name="sendCc">The CC recipient(s) of the email, specified as an array of string. Can be null or empty.</param>
    /// <param name="sendBcc">The BCC recipient(s) of the email, specified as an array of string. Can be null or empty.</param>
    /// <param name="body">The content of the email body. Can be plain text or HTML, depending on the <paramref name="useHtml"/> parameter.</param>
    /// <param name="useHtml">A boolean value indicating whether the email body should be formatted as HTML. If <see langword="true"/>, the
    /// <paramref name="body"/> is treated as HTML; otherwise, it is treated as plain text. Defaults to <see
    /// langword="false"/>.</param>
    /// <returns>An <see cref="Outlook.MailItem"/> representing the created email.</returns>
    /// <exception cref="McpException">Thrown if <paramref name="subject"/> or <paramref name="sendTo"/> is null or empty, or if an error occurs while
    /// creating the mail item.</exception>
    public Outlook.MailItem CreateMailItem(string subject, string[] sendTo, string[] sendCc, string[] sendBcc, string body, bool useHtml = false)
    {
        Outlook.MailItem mailItem;
        if (string.IsNullOrEmpty(subject))
        {
            throw new McpException($"Subject cannot be empty.");
        }
        if (sendTo == null || sendTo.Length == 0)
        {
            throw new McpException($"sendTo cannot be empty.");
        }
        try
        {
            mailItem = (Outlook.MailItem)Application.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = subject;
            mailItem.To = string.Join(";", sendTo);
            if (sendCc != null && sendCc.Length != 0)
            {
                mailItem.CC = string.Join(";", sendCc);
            }
            if (sendBcc != null && sendBcc.Length != 0)
            {
                mailItem.BCC = string.Join(";", sendBcc);
            }
            if (useHtml)
            {
                mailItem.HTMLBody = body;
            }
            else
            {
                mailItem.Body = body;
            }
        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to create mail item: {ex.Message}", ex);
        }
        RegisterComObject(mailItem);
        return mailItem;
    }

    /// <summary>
    /// Creates a new Outlook appointment item with the specified details.
    /// </summary>
    /// <remarks>This method creates an Outlook appointment item and sets its properties based on the provided
    /// parameters. The appointment is configured as a meeting, and the specified participants are added as
    /// recipients.</remarks>
    /// <param name="subject">The subject of the appointment. Cannot be null or empty.</param>
    /// <param name="participants">An array of email addresses representing the participants of the meeting. Cannot be null or empty.</param>
    /// <param name="startTime">The start time of the appointment.</param>
    /// <param name="duration">The duration of the appointment in minutes. Must be greater than 0.</param>
    /// <param name="location">The location of the appointment.</param>
    /// <param name="body">The body content of the appointment.</param>
    /// <returns>An <see cref="Outlook.AppointmentItem"/> representing the created appointment.</returns>
    /// <exception cref="McpException">Thrown if <paramref name="subject"/> is null or empty, <paramref name="participants"/> is null or empty,
    /// <paramref name="duration"/> is less than or equal to 0, or if an error occurs while creating the appointment.</exception>
    public Outlook.AppointmentItem CreateAppointmentItem(string subject, string[] participants, DateTime startTime, int duration, string? location = null, string? body = null)
    {
        Outlook.AppointmentItem appointmentItem;
        if (string.IsNullOrEmpty(subject))
        {
            throw new McpException($"Subject cannot be empty.");
        }
        if (duration <= 0)
        {
            throw new McpException($"Duration must be greater than 0.");
        }
        if (participants == null || participants.Length == 0)
        {
            throw new McpException($"Participants cannot be empty.");
        }
        try
        {
            appointmentItem = (Outlook.AppointmentItem)Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            appointmentItem.Subject = subject;
            var recipients = appointmentItem.Recipients;
            RegisterComObject(recipients);
            foreach (var p in participants)
            {
                var recipient = recipients.Add(p);
                RegisterComObject(recipient);
            }
            appointmentItem.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            appointmentItem.Start = startTime;
            appointmentItem.Duration = duration;
            if (!string.IsNullOrEmpty(location))
            {
                appointmentItem.Location = location;
            }
            if (!string.IsNullOrEmpty(body))
            {
                appointmentItem.Body = body;
            }

        }
        catch (Exception ex)
        {
            throw new McpException($"Failed to create appointment item: {ex.Message}", ex);
        }
        RegisterComObject(appointmentItem);
        return appointmentItem;
    }

}
