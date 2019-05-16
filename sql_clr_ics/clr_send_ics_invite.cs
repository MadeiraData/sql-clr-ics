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
        , SqlString smtp_server, SqlString username, SqlString password
        , SqlInt32 reminder_minutes)
    {
        // Put your code here
        #region validations
        // Validations
        StringBuilder sb_Errors = new StringBuilder();

        if (String.IsNullOrEmpty(from.Value)) sb_Errors.AppendLine("Missing parameter: from");
        if (String.IsNullOrEmpty(to.Value)) sb_Errors.AppendLine("Missing parameter: to");
        if (String.IsNullOrEmpty(subject.Value)) sb_Errors.AppendLine("Missing parameter: subject");
        if (String.IsNullOrEmpty(smtp_server.Value)) sb_Errors.AppendLine("Missing parameter: smtp_server");
        if (String.IsNullOrEmpty(username.Value)) sb_Errors.AppendLine("Missing parameter: username");
        if (String.IsNullOrEmpty(password.Value)) sb_Errors.AppendLine("Missing parameter: password");
        if (start_time.IsNull) sb_Errors.AppendLine("Missing parameter: start_time");
        if (end_time.IsNull) sb_Errors.AppendLine("Missing parameter: end_time");

        if (sb_Errors.Length > 0) throw new Exception("Unable to send mail due to validation error(s): " + sb_Errors);
        #endregion

        MailMessage msg = new MailMessage();
        //Now we have to set the value to Mail message properties

        //Note Please change it to correct mail-id to use this in your application
        msg.From = new MailAddress(from.Value);
        msg.To.Add(new MailAddress(to.Value));
        if (!String.IsNullOrEmpty(cc.Value)) msg.CC.Add(new MailAddress(cc.Value));
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

        if (!reminder_minutes.IsNull)
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
        smtpclient.Host = smtp_server.Value;

        smtpclient.Credentials = new NetworkCredential(username.Value, password.Value);

        System.Net.Mime.ContentType contype = new System.Net.Mime.ContentType("text/calendar");
        contype.Parameters.Add("method", "REQUEST");
        contype.Parameters.Add("name", "Meeting.ics");
        AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), contype);
        msg.AlternateViews.Add(avCal);
        msg.Headers.Add("Content-class", "urn:content-classes:calendarmessage");
        smtpclient.Send(msg);
    }
}
