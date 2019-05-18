using System;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Data.SqlTypes;

public partial class StoredProcedures
{
    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_send_ics_invite(
          SqlString from, SqlString to, SqlString cc, SqlString reply_to
        , SqlString subject, SqlString body, SqlString location
        , SqlDateTime start_time_utc, SqlDateTime end_time_utc, SqlDateTime timestamp_utc
        , SqlString smtp_server, SqlInt32 port, SqlBoolean use_ssl, SqlString username, SqlString password
        , SqlBoolean use_reminder, SqlInt32 reminder_minutes
        , SqlGuid cancel_event_identifier, out SqlGuid event_identifier
        , SqlBoolean suppress_info_messages
        )
    {
        #region validations

        StringBuilder sb_Errors = new StringBuilder();

        if (from.IsNull || String.IsNullOrEmpty(from.Value)) sb_Errors.AppendLine("Missing parameter: from");
        if (to.IsNull || String.IsNullOrEmpty(to.Value)) sb_Errors.AppendLine("Missing parameter: to");
        if (subject.IsNull || String.IsNullOrEmpty(subject.Value)) sb_Errors.AppendLine("Missing parameter: subject");

        if (sb_Errors.Length > 0) throw new Exception("Unable to send mail due to validation error(s): " + sb_Errors);

        #endregion validations

        #region default values initialization

        if (start_time_utc.IsNull) start_time_utc = DateTime.Now.AddMinutes(+300);
        if (end_time_utc.IsNull) end_time_utc = start_time_utc.Value.AddMinutes(+60);
        if (reminder_minutes.IsNull) reminder_minutes = 15;
        if (use_reminder.IsNull) use_reminder = true;
        if (suppress_info_messages.IsNull) suppress_info_messages = false;

        if (smtp_server.IsNull || String.IsNullOrEmpty(smtp_server.Value)) smtp_server = "localhost";
        if (port.IsNull) port = 25;
        if (use_ssl.IsNull) use_ssl = false;
        if (timestamp_utc.IsNull) timestamp_utc = DateTime.UtcNow;

        bool useDefaultCredentials = false;
        ICredentialsByHost credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

        if (username.IsNull || String.IsNullOrEmpty(username.Value))
        {
            useDefaultCredentials = true;
        }
        else
        {
            if (password.IsNull || String.IsNullOrEmpty(password.Value)) password = "";
            credentials = new NetworkCredential(username.Value, password.Value);
        }

        if (!cancel_event_identifier.IsNull)
            event_identifier = cancel_event_identifier.Value;
        else
            event_identifier = Guid.NewGuid();

        #endregion default values initialization

        #region initialize MailMessage and recipients

        MailMessage msg = new MailMessage();
        msg.Subject = subject.Value;
        msg.Body = body.Value;

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
            if (!reply_to.IsNull && !String.IsNullOrEmpty(reply_to.Value)) msg.ReplyToList.Add(reply_to.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("ReplyTo address is invalid: " + e.Message);
        }

        try
        {
            msg.To.Add(to.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("To address is invalid: " + e.Message);
        }

        try
        {
            if (!cc.IsNull && !String.IsNullOrEmpty(cc.Value)) msg.CC.Add(cc.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("CC address is invalid: " + e.Message);
        }

        #endregion initialize MailMessage and recipients

        #region construct ICS file contents

        StringBuilder str = new StringBuilder();
        str.AppendLine("BEGIN:VCALENDAR");
        str.AppendLine("PRODID:-//Schedule a Meeting");
        str.AppendLine("VERSION:2.0");

        if (cancel_event_identifier.IsNull)
        {
            str.AppendLine("METHOD:REQUEST");
            str.AppendLine("SEQUENCE:1");
        }
        else
        {
            str.AppendLine("METHOD:CANCEL");
            str.AppendLine("SEQUENCE:2");
        }

        str.AppendLine("BEGIN:VEVENT");
        str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", start_time_utc.Value));
        str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", timestamp_utc.Value));
        str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", end_time_utc.Value));
        if (!location.IsNull) str.AppendLine("LOCATION: " + location.Value);
        str.AppendLine(string.Format("UID:{0}", event_identifier.Value));
        str.AppendLine(string.Format("DESCRIPTION:{0}", msg.Body));
        str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", msg.Body));
        str.AppendLine(string.Format("SUMMARY:{0}", msg.Subject));
        str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", msg.From.Address));

        foreach (MailAddress addr in msg.To)
        {
            str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=\"{0}\";X-NUM-UESTS=0:mailto:{1}", addr.DisplayName, addr.Address));
        }

        if (use_reminder && cancel_event_identifier.IsNull)
        {
            str.AppendLine("BEGIN:VALARM");
            str.AppendLine(string.Format("TRIGGER:-PT{0}M", reminder_minutes.Value));
            str.AppendLine("ACTION:DISPLAY");
            str.AppendLine("DESCRIPTION:Reminder");
            str.AppendLine("END:VALARM");
        }

        str.AppendLine("END:VEVENT");
        str.AppendLine("END:VCALENDAR");

        #endregion construct ICS file contents

        #region initialize and configure SmtpClient

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

        #endregion initialize and configure SmtpClient

        try
        {
            smtpclient.Send(msg);
            if (!suppress_info_messages)
                Microsoft.SqlServer.Server.SqlContext.Pipe.Send(string.Format("Mail Sent. Event Identifier: {0}", event_identifier.Value));
        }
        catch (Exception e)
        {
            throw new Exception("Error sending mail: " + e.Message);
        }
    }
}