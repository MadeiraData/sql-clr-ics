using System;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using Microsoft.SqlServer.Server;

public partial class StoredProcedures
{
    [Microsoft.SqlServer.Server.SqlProcedure]
    public static void clr_send_ics_invite(
          SqlString profile_name
        , SqlString recipients, SqlString copy_recipients, SqlString blind_copy_recipients
        , SqlString from_address, SqlString reply_to
        , SqlString subject, SqlString body
        , SqlString importance, SqlString sensitivity, SqlString file_attachments
        , SqlString location, SqlDateTime start_time_utc, SqlDateTime end_time_utc, SqlDateTime timestamp_utc
        , SqlString method, SqlInt32 sequence, SqlString prod_id
        , SqlBoolean use_reminder, SqlInt32 reminder_minutes, SqlBoolean require_rsvp
        , SqlString smtp_servername, SqlInt32 port, SqlBoolean enable_ssl
        , SqlString username, SqlString password
        , SqlBoolean suppress_info_messages
        , ref SqlGuid event_identifier
        )
    {
        #region get missing info from sysmail profile

        if (from_address.IsNull)
        {
            SqlConnection con = new SqlConnection("context connection=true"); // using existing CLR context connection
            SqlCommand cmd = con.CreateCommand();
            con.Open();

            if (profile_name.IsNull)
            {
                cmd.CommandText = @"SELECT p.name
FROM [msdb].[dbo].[sysmail_principalprofile] AS pp
INNER JOIN [msdb].[dbo].[sysmail_profile] AS p
ON pp.profile_id = p.profile_id
WHERE pp.is_default = 1";

                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    if (!rdr.HasRows)
                    {
                        rdr.Close();
                        con.Close();
                        throw new Exception("profile_name not specified and no default profile found");
                    } else
                    {
                        profile_name = rdr.GetSqlString(0);
                    }
                    rdr.Close();
                }
            }


            con.Close();
        }

        #endregion get missing info from sysmail profile

        #region validations

        StringBuilder sb_Errors = new StringBuilder();

        if (from_address.IsNull || String.IsNullOrEmpty(from_address.Value)) sb_Errors.AppendLine("Missing parameter: from");
        if (recipients.IsNull || String.IsNullOrEmpty(recipients.Value)) sb_Errors.AppendLine("Missing parameter: to");
        if (subject.IsNull || String.IsNullOrEmpty(subject.Value)) sb_Errors.AppendLine("Missing parameter: subject");

        if (sb_Errors.Length > 0) throw new Exception("Unable to send mail due to validation error(s): " + sb_Errors);

        #endregion validations

        #region default values initialization

        if (start_time_utc.IsNull) start_time_utc = DateTime.Now.AddMinutes(+300);
        if (end_time_utc.IsNull) end_time_utc = start_time_utc.Value.AddMinutes(+60);
        if (reminder_minutes.IsNull) reminder_minutes = 15;
        if (use_reminder.IsNull) use_reminder = true;
        if (suppress_info_messages.IsNull) suppress_info_messages = false;
        if (prod_id.IsNull || String.IsNullOrEmpty(prod_id.Value)) prod_id = "-//Schedule a Meeting";

        if (smtp_servername.IsNull || String.IsNullOrEmpty(smtp_servername.Value)) smtp_servername = "localhost";
        if (port.IsNull) port = 25;
        if (enable_ssl.IsNull) enable_ssl = false;
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

        #endregion default values initialization

#region initialize MailMessage and recipients

        MailMessage msg = new MailMessage();
        msg.Subject = subject.Value;
        msg.Body = body.Value;

        try
        {
            msg.From = new MailAddress(from_address.Value);
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
            msg.To.Add(recipients.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("To address is invalid: " + e.Message);
        }

        try
        {
            if (!copy_recipients.IsNull && !String.IsNullOrEmpty(copy_recipients.Value)) msg.CC.Add(copy_recipients.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("CC address is invalid: " + e.Message);
        }

#endregion initialize MailMessage and recipients

#region construct ICS file contents

        StringBuilder str = new StringBuilder();
        str.AppendLine("BEGIN:VCALENDAR");
        str.AppendLine(string.Format("PRODID:{0}", prod_id.Value));
        str.AppendLine("VERSION:2.0");

        if (method == "request")
        {
            str.AppendLine("METHOD:REQUEST");
            str.AppendLine("SEQUENCE:0");
        }
        else if (method == "cancel")
        {
            str.AppendLine("METHOD:CANCEL");
            str.AppendLine("SEQUENCE:1");
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

        string rsvp_string = (require_rsvp.Value ? "PARTSTAT=NEEDS-ACTION;RSVP=TRUE" : "PARTSTAT=ACCEPTED;RSVP=FALSE");
        bool organizer_in_recipients = false;

        foreach (MailAddress addr in msg.To)
        {
            if (addr.Address == msg.From.Address) organizer_in_recipients = true;
            str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;{2};CN=\"{0}\";X-NUM-GUESTS=0:mailto:{1}", addr.DisplayName, addr.Address, rsvp_string));
        }

        if (!organizer_in_recipients) str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=NON-PARTICIPANT;PARTSTAT=ACCEPTED;RSVP=FALSE;CN=\"{0}\";X-NUM-GUESTS=0:mailto:{1}", msg.From.DisplayName, msg.From.Address));

        if (use_reminder && method.IsNull)
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
            smtpclient.Host = smtp_servername.Value;
            smtpclient.Port = port.Value;
            smtpclient.UseDefaultCredentials = useDefaultCredentials;
            smtpclient.EnableSsl = enable_ssl.Value;
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
                SqlContext.Pipe.Send(string.Format("Mail Sent. Event Identifier: {0}", event_identifier.Value));
        }
        catch (Exception e)
        {
            throw new Exception("Error sending mail: " + e.Message);
        }
    }
}