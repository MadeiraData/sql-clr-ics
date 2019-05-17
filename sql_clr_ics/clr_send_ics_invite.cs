using System;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Data.SqlTypes;

public partial class StoredProcedures
{
    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_send_ics_invite (
          SqlString from, SqlString to, SqlString cc
        , SqlString subject, SqlString body, SqlString location, SqlDateTime start_time, SqlDateTime end_time
        , SqlString smtp_server, SqlInt32 port, SqlBoolean use_ssl, SqlString username, SqlString password
        , SqlBoolean use_reminder, SqlInt32 reminder_minutes)
    {
        // Validations
        #region validations

        StringBuilder sb_Errors = new StringBuilder();

        if (String.IsNullOrEmpty(from.Value)) sb_Errors.AppendLine("Missing parameter: from");
        if (String.IsNullOrEmpty(to.Value)) sb_Errors.AppendLine("Missing parameter: to");
        if (String.IsNullOrEmpty(subject.Value)) sb_Errors.AppendLine("Missing parameter: subject");

        if (sb_Errors.Length > 0) throw new Exception("Unable to send mail due to validation error(s): " + sb_Errors);

        #endregion

        // Default Values Initialization
        #region default values

        if (start_time.IsNull) start_time = DateTime.Now.AddMinutes(+300);
        if (end_time.IsNull) end_time = start_time.Value.AddMinutes(+60);
        if (reminder_minutes.IsNull) reminder_minutes = 15;
        if (use_reminder.IsNull) use_reminder = true;

        if (String.IsNullOrEmpty(smtp_server.Value)) smtp_server = "localhost";
        if (port.IsNull) port = 25;
        if (use_ssl.IsNull) use_ssl = false;

        bool useDefaultCredentials = false;
        ICredentialsByHost credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

        if (String.IsNullOrEmpty(username.Value))
        {
            useDefaultCredentials = true;
        }
        else
        {
            if (String.IsNullOrEmpty(password.Value)) password = "";
            credentials = new NetworkCredential(username.Value, password.Value);
        }

        #endregion

        MailMessage msg = new MailMessage();

        //Now we have to set the value to Mail message properties

        try
        {
            msg.From = new MailAddress(from.Value);
        }
        catch (Exception e)
        {
            throw new Exception("From address is invalid: " + e.Message);
        }

        try
        {
            msg.To.Add(new MailAddress(to.Value));
        }
        catch (Exception e)
        {
            throw new Exception("To address is invalid: " + e.Message);
        }


        try
        {
            if (!cc.IsNull && !String.IsNullOrEmpty(cc.Value)) msg.CC.Add(new MailAddress(cc.Value));
        }
        catch (Exception e)
        {
            throw new Exception("CC address is invalid: " + e.Message);
        }

        msg.Subject = subject.Value;
        msg.Body = body.Value;

        // Now Contruct the ICS file using string builder
        StringBuilder str = new StringBuilder();
        str.AppendLine("BEGIN:VCALENDAR");
        str.AppendLine("PRODID:-//Schedule a Meeting");
        str.AppendLine("VERSION:2.0");
        str.AppendLine("METHOD:REQUEST");
        str.AppendLine("BEGIN:VEVENT");
        str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", start_time.Value));
        str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
        str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", end_time.Value));
        if (!location.IsNull) str.AppendLine("LOCATION: " + location.Value);
        str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
        str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
        str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
        str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
        str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

        str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", msg.To[0].DisplayName, msg.To[0].Address));

        if (use_reminder)
        {
            str.AppendLine("BEGIN:VALARM");
            str.AppendLine(string.Format("TRIGGER:-PT{0}M", reminder_minutes.Value));
            str.AppendLine("ACTION:DISPLAY");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("END:VALARM");
        }

        str.AppendLine("END:VEVENT");
        str.AppendLine("END:VCALENDAR");

        //Now sending a mail with attachment ICS file.

        System.Net.Mail.SmtpClient smtpclient = new System.Net.Mail.SmtpClient();

        try
        {
            smtpclient.Host = smtp_server.Value;
            smtpclient.Port = port.Value;
            smtpclient.UseDefaultCredentials = useDefaultCredentials;
            smtpclient.EnableSsl = use_ssl.Value;
            smtpclient.Credentials = credentials;
            System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
            contype.Parameters.Add("method", "REQUEST");
            contype.Parameters.Add("name", "Meeting.ics");

            AlternateView HTML = AlternateView.CreateAlternateViewFromString(body.Value, new System.Net.Mime.ContentType("text/html"));
            msg.AlternateViews.Add(HTML);
            AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
            msg.AlternateViews.Add(avCal);

            msg.Headers.Add("Content-class", "urn:content-classes:calendarmessage");
        }
        catch (Exception e)
        {
            throw new Exception("SMTP Client Configuration Error: " + e.Message);
        }

        try
        {
            smtpclient.Send(msg);
        }
        catch (Exception e)
        {
            throw new Exception("Error while trying to send mail: " + e.Message);
        }
    }
}
