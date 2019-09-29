using System;
using System.Text;
using System.Net;
using System.Net.Mail;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using Microsoft.SqlServer.Server;

public partial class StoredProcedures
{
    #region Enums and Constants
    private enum iCalMethods
    {
        PUBLISH,
        REQUEST,
        REPLY,
        CANCEL,
        ADD,
        REFRESH,
        COUNTER,
        DECLINECOUNTER
    }

    private enum iCalClass
    {
        PUBLIC,
        PRIVATE,
        CONFIDENTIAL
    }

    private static string[] iCalRoles = { "REQ-PARTICIPANT", "OPT-PARTICIPANT", "NON-PARTICIPANT", "CHAIR" };
    #endregion Enums and Constants

    [SqlProcedure]
    public static void clr_send_ics_invite(
          SqlString profile_name
        , SqlString recipients, SqlString copy_recipients, SqlString blind_copy_recipients
        , SqlString from_address, SqlString reply_to
        , SqlString subject, SqlString body, SqlString body_format
        , SqlString importance, SqlString sensitivity, SqlString file_attachments
        , SqlString location, SqlDateTime start_time_utc, SqlDateTime end_time_utc, SqlDateTime timestamp_utc
        , SqlString method, SqlInt32 sequence, SqlString prod_id
        , SqlBoolean use_reminder, SqlInt32 reminder_minutes, SqlBoolean require_rsvp
        , SqlString recipients_role, SqlString copy_recipients_role, SqlString blind_copy_recipients_role
        , SqlString smtp_servername, SqlInt32 port, SqlBoolean enable_ssl
        , SqlBoolean use_default_credentials, SqlString username, SqlString password
        , SqlBoolean suppress_info_messages
        , ref SqlString event_identifier, ref SqlString ics_contents
        )
    {
        #region local variable declaration

        ICredentialsByHost credentials = CredentialCache.DefaultNetworkCredentials;
        MailPriority mailPriority;

        #endregion local variable declaration

        #region get missing info from sysmail profile

        string currentPhase = "";
        try
        {
        if (from_address.IsNull || (username.IsNull && use_default_credentials.IsNull) || !profile_name.IsNull)
        {
            currentPhase = "Creating SqlConnection";
            SqlConnection con = new SqlConnection("context connection=true"); // using existing CLR context connection
            currentPhase = "Creating SqlCommand";
            SqlCommand cmd = con.CreateCommand();
            currentPhase = "Opening SqlConnection";
            con.Open();

            if (profile_name.IsNull)
            {
                cmd.CommandText = @"SELECT p.name
FROM [msdb].[dbo].[sysmail_principalprofile] AS pp
INNER JOIN [msdb].[dbo].[sysmail_profile] AS p
ON pp.profile_id = p.profile_id
WHERE pp.is_default = 1";

                currentPhase = "Getting default DBMail profile";
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {
                    if (!rdr.HasRows)
                    {
                        rdr.Close();
                        con.Close();
                        throw new Exception("profile_name not specified and no default profile found");
                    }
                    else
                    {
                        profile_name = rdr.GetSqlString(0);
                    }
                    rdr.Close();
                }
            }

            cmd.CommandText = @"SELECT TOP 1 a.email_address, a.replyto_address, s.servername, s.port, s.enable_ssl, s.use_default_credentials, s.username
FROM [msdb].[dbo].[sysmail_profile] AS p
INNER JOIN [msdb].[dbo].[sysmail_profileaccount] AS pa
ON p.profile_id = pa.profile_id
AND pa.sequence_number >= @Seq
INNER JOIN [msdb].[dbo].[sysmail_account] AS a
ON pa.account_id = a.account_id
INNER JOIN [msdb].[dbo].[sysmail_server] AS s
ON p.profile_id = s.account_id
WHERE p.name = @Profile
ORDER BY pa.sequence_number ASC";

            cmd.Parameters.AddWithValue("@Seq", 1);
            cmd.Parameters.AddWithValue("@Profile", profile_name.Value);

            currentPhase = string.Format("Getting profile settings ({0})", profile_name.Value);
            using (SqlDataReader rdr = cmd.ExecuteReader())
            {
                if (!rdr.HasRows)
                {
                    rdr.Close();
                    con.Close();
                    throw new Exception(string.Format("profile_name '{0}' not found", profile_name.Value));
                }
                else
                {
                    rdr.Read();
                    if (from_address.IsNull) from_address = rdr.GetSqlString(0);
                    if (reply_to.IsNull) reply_to = rdr.GetSqlString(1);
                    if (smtp_servername.IsNull) smtp_servername = rdr.GetSqlString(2);
                    if (port.IsNull) port = rdr.GetSqlInt32(3);
                    if (enable_ssl.IsNull) enable_ssl = rdr.GetSqlBoolean(4);
                    if (use_default_credentials.IsNull) use_default_credentials = rdr.GetSqlBoolean(5).Value;
                    if (username.IsNull && !use_default_credentials.Value) username = rdr.GetSqlString(6);
                }
                rdr.Close();
            }
            con.Close();
            }

        }
        catch (Exception ex)
        {
            throw new Exception(string.Format("Error while {0}: {1}", currentPhase, ex.Message), ex);
        }

        #endregion get missing info from sysmail profile

        #region default values initialization

        if (subject.IsNull || string.IsNullOrEmpty(subject.Value)) subject = "SQL Server Meeting";
        if (body_format.IsNull || string.IsNullOrEmpty(body_format.Value)) body_format = "TEXT";
        if (timestamp_utc.IsNull) timestamp_utc = DateTime.UtcNow;
        if (start_time_utc.IsNull) start_time_utc = timestamp_utc.Value.AddMinutes(+300);
        if (end_time_utc.IsNull) end_time_utc = start_time_utc.Value.AddMinutes(+60);
        if (reminder_minutes.IsNull) reminder_minutes = 15;
        if (use_reminder.IsNull) use_reminder = true;
        if (suppress_info_messages.IsNull) suppress_info_messages = false;
        if (prod_id.IsNull || string.IsNullOrEmpty(prod_id.Value)) prod_id = "Schedule a Meeting";
        if (importance.IsNull || string.IsNullOrEmpty(importance.Value)) importance = "Normal";
        if (sensitivity.IsNull || string.IsNullOrEmpty(sensitivity.Value)) sensitivity = "Public";

        if (smtp_servername.IsNull || string.IsNullOrEmpty(smtp_servername.Value)) smtp_servername = "localhost";
        if (port.IsNull) port = 25;
        if (enable_ssl.IsNull) enable_ssl = false;

        if (username.IsNull || string.IsNullOrEmpty(username.Value))
        {
            use_default_credentials = true;
        }
        else
        {
            if (password.IsNull || string.IsNullOrEmpty(password.Value)) password = "";
            credentials = new NetworkCredential(username.Value, password.Value);
        }

        if (recipients_role.IsNull) recipients_role = "REQ-PARTICIPANT";
        if (copy_recipients_role.IsNull) copy_recipients_role = "OPT-PARTICIPANT";
        if (blind_copy_recipients_role.IsNull) blind_copy_recipients_role = "NON-PARTICIPANT";

        if (method.IsNull) method = "REQUEST";
        if (sequence.IsNull) sequence = (method.Value == "CANCEL" ? 1 : 0);
        if (event_identifier.IsNull) event_identifier = Guid.NewGuid().ToString();

        #endregion default values initialization

        #region validations

        StringBuilder sb_Errors = new StringBuilder();

        if (from_address.IsNull || string.IsNullOrEmpty(from_address.Value)) sb_Errors.AppendLine("Missing sender: Please specify @from_address");
        if (
                (recipients.IsNull || string.IsNullOrEmpty(recipients.Value))
            && (copy_recipients.IsNull || string.IsNullOrEmpty(copy_recipients.Value))
            && (blind_copy_recipients.IsNull || string.IsNullOrEmpty(blind_copy_recipients.Value))
           )
            sb_Errors.AppendLine("Missing recipients: Please specify either @recipients, @copy_recipients or @blind_copy_recipients");

        if (body_format.Value != "HTML" && body_format.Value != "TEXT") sb_Errors.AppendLine(string.Format("@body_format {0} is invalid. Valid values: TEXT, HTML", body_format.Value));
        if (!Enum.TryParse(method.Value, true, out iCalMethods method_enumvalue)) sb_Errors.AppendLine(string.Format("@method {0} is invalid. Valid values: {1}", method.Value, Enum.GetNames(typeof(iCalMethods)).ToString().ToUpper()));
        if (!Enum.TryParse(sensitivity.Value, true, out iCalClass sensitivity_enumvalue)) sb_Errors.AppendLine(string.Format("sensitivity {0} is invalid. Valid values: {1}", sensitivity.Value, Enum.GetNames(typeof(iCalClass)).ToString().ToUpper()));
        if (!Enum.TryParse(importance.Value, true, out mailPriority)) sb_Errors.AppendLine(string.Format("@importance {0} is invalid. Valid values: {1}", importance.Value, Enum.GetNames(typeof(MailPriority)).ToString().ToUpper()));

        bool recipient_role_found = false;
        bool copy_recipient_role_found = false;
        bool blind_copy_recipient_role_found = false;

        foreach (var item in iCalRoles)
        {
            if (!recipients.IsNull && item == recipients_role.Value)
            {
                recipient_role_found = true;
            }
            if (!copy_recipients.IsNull && item == copy_recipients_role.Value)
            {
                copy_recipient_role_found = true;
            }
            if (!blind_copy_recipients.IsNull && item == blind_copy_recipients_role.Value)
            {
                blind_copy_recipient_role_found = true;
            }
            if (recipient_role_found && copy_recipient_role_found && blind_copy_recipient_role_found)
            {
                break;
            }
        }

        if (!recipients.IsNull && !recipient_role_found) sb_Errors.AppendLine(string.Format("@recipients_role {0} is invalid. Valid values: {1}", recipients_role.Value, iCalRoles.ToString()));
        if (!copy_recipients.IsNull && !copy_recipient_role_found) sb_Errors.AppendLine(string.Format("@copy_recipients_role {0} is invalid. Valid values: {1}", copy_recipients_role.Value, iCalRoles.ToString()));
        if (!blind_copy_recipients.IsNull && !blind_copy_recipient_role_found) sb_Errors.AppendLine(string.Format("@blind_copy_recipients_role {0} is invalid. Valid values: {1}", blind_copy_recipients_role.Value, iCalRoles.ToString()));

        if (sb_Errors.Length > 0) throw new Exception("Unable to send mail due to validation error(s): " + sb_Errors);

        #endregion validations

        #region initialize MailMessage and recipients

        MailMessage msg = new MailMessage();
        msg.Subject = subject.Value;
        msg.Body = body.Value;
        msg.Priority = mailPriority;
        msg.IsBodyHtml = (body_format.Value == "HTML");

        if (!file_attachments.IsNull)
        {
            foreach (var att_path in file_attachments.Value.Split(';'))
            {
                try
                {
                    msg.Attachments.Add(new Attachment(att_path));
                }
                catch (Exception e)
                {
                    throw new Exception(string.Format("Unable to attach '{0}': {1}", att_path, e.Message), e);
                }
            }
        }

        try
        {
            msg.From = new MailAddress(from_address.Value);
        }
        catch (Exception e)
        {
            throw new Exception("@from_address is invalid: " + e.Message, e);
        }

        try
        {
            if (!reply_to.IsNull && !string.IsNullOrEmpty(reply_to.Value)) msg.ReplyToList.Add(reply_to.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("@reply_to is invalid: " + e.Message, e);
        }

        try
        {
            msg.To.Add(recipients.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("@recipients is invalid: " + e.Message, e);
        }

        try
        {
            if (!copy_recipients.IsNull && !string.IsNullOrEmpty(copy_recipients.Value)) msg.CC.Add(copy_recipients.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("@copy_recipients is invalid: " + e.Message, e);
        }

        try
        {
            if (!blind_copy_recipients.IsNull && !string.IsNullOrEmpty(blind_copy_recipients.Value)) msg.Bcc.Add(blind_copy_recipients.Value.Replace(';', ','));
        }
        catch (Exception e)
        {
            throw new Exception("@blind_copy_recipients is invalid: " + e.Message, e);
        }

        #endregion initialize MailMessage and recipients

        #region construct ICS file contents

        if (ics_contents.IsNull)
        {
            StringBuilder ics_contents_str = new StringBuilder();
            ics_contents_str.AppendLine("BEGIN:VCALENDAR");
            ics_contents_str.AppendLine(string.Format("PRODID:-//{0}", prod_id.Value));
            ics_contents_str.AppendLine("VERSION:2.0");
            ics_contents_str.AppendLine(string.Format("METHOD:{0}", method.Value.ToUpper()));

            ics_contents_str.AppendLine("BEGIN:VEVENT");
            ics_contents_str.AppendLine(string.Format("STATUS:{0}", (method.Value == "CANCEL") ? "CANCELLED" : "CONFIRMED"));
            ics_contents_str.AppendLine("TRANSP:OPAQUE");
            ics_contents_str.AppendLine(string.Format("SEQUENCE:{0}", sequence.Value));
            ics_contents_str.AppendLine(string.Format("X-MICROSOFT-CDO-APPT-SEQUENCE:{0}", sequence.Value));

            ics_contents_str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", start_time_utc.Value));
            ics_contents_str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", timestamp_utc.Value));
            ics_contents_str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", end_time_utc.Value));
            if (!location.IsNull) ics_contents_str.AppendLine("LOCATION: " + location.Value);
            ics_contents_str.AppendLine(string.Format("UID:{0}", event_identifier.Value));
            ics_contents_str.AppendLine(string.Format("DESCRIPTION:{0}", body.Value));

            if (method.Value != "CANCEL")
                ics_contents_str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE={0}:{1}", body_format.Value == "HTML" ? "text/html" : "text/plain", body.Value));

            ics_contents_str.AppendLine(string.Format("SUMMARY:{0}", subject.Value));
            ics_contents_str.AppendLine(string.Format("ORGANIZER:{0}", msg.From.Address));
            //ics_contents_str.AppendLine(string.Format("ORGANIZER:SENT-BY=\"mailto:{0}\";MAILTO:{0}", msg.From.Address));
            ics_contents_str.AppendLine(string.Format("CLASS:{0}", sensitivity.Value.ToUpper()));

            switch (mailPriority)
            {
                case MailPriority.Normal:
                    ics_contents_str.AppendLine("PRIORITY:5");
                    break;
                case MailPriority.Low:
                    ics_contents_str.AppendLine("PRIORITY:9");
                    break;
                case MailPriority.High:
                    ics_contents_str.AppendLine("PRIORITY:1");
                    break;
                default:
                    break;
            }

            string rsvp_string = (require_rsvp.Value ? "PARTSTAT=NEEDS-ACTION;RSVP=TRUE" : "PARTSTAT=ACCEPTED;RSVP=FALSE");
            bool organizer_in_recipients = false;
            string lineSeparator = " "; // Environment.NewLine;

            foreach (MailAddress addr in msg.To)
            {
                if (addr.Address == msg.From.Address) organizer_in_recipients = true;
                ics_contents_str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE={3};{4}{2};CN=\"{0}\";{4}X-NUM-GUESTS=0:mailto:{1}", addr.DisplayName, addr.Address, rsvp_string, recipients_role.Value.ToUpper(), lineSeparator));
            }

            foreach (MailAddress addr in msg.CC)
            {
                if (addr.Address == msg.From.Address) organizer_in_recipients = true;
                ics_contents_str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE={3};{4}{2};CN=\"{0}\";{4}X-NUM-GUESTS=0:mailto:{1}", addr.DisplayName, addr.Address, rsvp_string, copy_recipients_role.Value.ToUpper(), lineSeparator));
            }

            foreach (MailAddress addr in msg.Bcc)
            {
                if (addr.Address == msg.From.Address) organizer_in_recipients = true;
                ics_contents_str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE={3};{4}{2};CN=\"{0}\";{4}X-NUM-GUESTS=0:mailto:{1}", addr.DisplayName, addr.Address, rsvp_string, blind_copy_recipients_role.Value.ToUpper(), lineSeparator));
            }

            if (!organizer_in_recipients)
            {
                msg.Bcc.Add(msg.From);
                ics_contents_str.AppendLine(string.Format("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=NON-PARTICIPANT;{2}PARTSTAT=ACCEPTED;RSVP=FALSE;CN=\"{0}\";{2}X-NUM-GUESTS=0:mailto:{1}", msg.From.DisplayName, msg.From.Address, lineSeparator));
            }

            if (use_reminder && method.Value != "CANCEL")
            {
                ics_contents_str.AppendLine("BEGIN:VALARM");
                ics_contents_str.AppendLine(string.Format("TRIGGER:-PT{0}M", reminder_minutes.Value));
                ics_contents_str.AppendLine("ACTION:DISPLAY");
                ics_contents_str.AppendLine("DESCRIPTION:REMINDER");
                ics_contents_str.AppendLine("END:VALARM");
            }

            ics_contents_str.AppendLine("END:VEVENT");
            ics_contents_str.AppendLine("END:VCALENDAR");

            ics_contents = ics_contents_str.ToString();
        }

        #endregion construct ICS file contents

        #region initialize and configure SmtpClient

        SmtpClient smtpclient = new SmtpClient();

        try
        {
            smtpclient.Host = smtp_servername.Value;
            smtpclient.Port = port.Value;
            smtpclient.UseDefaultCredentials = use_default_credentials.Value;
            smtpclient.EnableSsl = enable_ssl.Value;
            smtpclient.Credentials = credentials;
            smtpclient.DeliveryMethod = SmtpDeliveryMethod.Network;
            System.Net.Mime.ContentType calendar_contype = new System.Net.Mime.ContentType("text/calendar;charset=UTF-8");
            calendar_contype.Parameters.Add("method", "REQUEST");
            calendar_contype.Parameters.Add("name", "Meeting.ics");

            AlternateView avBody = AlternateView.CreateAlternateViewFromString(body.Value, new System.Net.Mime.ContentType(body_format.Value == "HTML" ? "text/html;charset=UTF-8" : "text/plain"));
            msg.AlternateViews.Add(avBody);

            AlternateView avCal = AlternateView.CreateAlternateViewFromString(ics_contents.Value, calendar_contype);
            //avCal.TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
            msg.AlternateViews.Add(avCal);

            msg.Headers.Add("Content-class", "urn:content-classes:calendarmessage");
        }
        catch (Exception e)
        {
            throw new Exception("SMTP Client Configuration Error: " + e.Message, e);
        }

        #endregion initialize and configure SmtpClient

        #region send mail

        try
        {
            //if (!suppress_info_messages)
            //    SqlContext.Pipe.Send(string.Format("Sending message to: {0}{1}{2}", msg.To.ToString(), Environment.NewLine, msg.Body));

            smtpclient.Send(msg);

            if (!suppress_info_messages)
                SqlContext.Pipe.Send(string.Format("Mail Sent. Event Identifier: {0}", event_identifier.Value));
        }
        catch (Exception e)
        {
            throw new Exception("Error sending mail: " + e.Message, e);
        }

        #endregion send mail
    }
}