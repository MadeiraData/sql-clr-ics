using System;
using System.IO;
using System.Text;
using System.Net;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Text.RegularExpressions;
using Microsoft.SqlServer.Server;

/*
    clr_send_graph_event
    ---------------------------------------------------
    A Microsoft Graph API equivalent of clr_send_ics_invite.

    Instead of hand-building an ICS (VCALENDAR/VEVENT) document and sending it
    over SMTP, this procedure creates / updates / cancels a calendar event
    directly in a Microsoft 365 mailbox via the Microsoft Graph REST API.
    Graph takes care of dispatching the meeting invitations (and any
    cancellation notices) to the attendees on the organizer's behalf.

    Authentication uses the OAuth2 client-credentials (app-only) flow, which is
    the appropriate model for unattended SQL Server automation. Alternatively,
    a caller may supply a pre-acquired bearer token via @access_token (e.g. from
    an external auth broker), in which case the client credentials are ignored.

    The assembly intentionally has no external dependencies (no Microsoft.Graph
    SDK, no JSON library) so that SQL-CLR deployment stays as simple as the
    original ICS assembly: everything is done with System.Net.HttpWebRequest and
    lightweight, hand-written JSON building/extraction.
*/
public partial class StoredProcedures
{
    #region Graph enums, constants and private members

    private enum GraphMethods
    {
        REQUEST,    // create a new event and send invitations
        UPDATE,     // patch an existing event (identified by @event_identifier)
        CANCEL      // cancel an existing event and notify attendees
    }

    private static bool m_graph_suppress_info_messages = false;

    private const string DefaultAuthorityUrl = "https://login.microsoftonline.com";
    private const string DefaultGraphUrl = "https://graph.microsoft.com";

    #endregion Graph enums, constants and private members

    [SqlProcedure]
    public static void clr_send_graph_event(
          SqlString tenant_id, SqlString client_id, SqlString client_secret
        , SqlString access_token
        , SqlString authority_url, SqlString graph_url
        , SqlString organizer
        , SqlString recipients, SqlString optional_recipients, SqlString resource_recipients
        , SqlString subject, SqlString body, SqlString body_format
        , SqlString importance, SqlString sensitivity
        , SqlString location
        , SqlDateTime start_time_utc, SqlDateTime end_time_utc
        , SqlString method
        , SqlBoolean use_reminder, SqlInt32 reminder_minutes
        , SqlBoolean require_rsvp
        , SqlBoolean create_teams_meeting
        , SqlBoolean all_day_event
        , SqlString cancellation_comment
        , SqlBoolean suppress_info_messages
        , ref SqlString event_identifier, ref SqlString response_content
        )
    {
        string currentPhase = "";
        try
        {
            #region default values initialization

            currentPhase = "Initializing parameter defaults";

            if (subject.IsNull || string.IsNullOrEmpty(subject.Value)) subject = "SQL Server Meeting";
            if (body_format.IsNull || string.IsNullOrEmpty(body_format.Value)) body_format = "TEXT";
            if (importance.IsNull || string.IsNullOrEmpty(importance.Value)) importance = "Normal";
            if (sensitivity.IsNull || string.IsNullOrEmpty(sensitivity.Value)) sensitivity = "Public";
            if (method.IsNull || string.IsNullOrEmpty(method.Value)) method = "REQUEST";
            if (use_reminder.IsNull) use_reminder = true;
            if (reminder_minutes.IsNull) reminder_minutes = 15;
            if (require_rsvp.IsNull) require_rsvp = true;
            if (create_teams_meeting.IsNull) create_teams_meeting = false;
            if (all_day_event.IsNull) all_day_event = false;
            if (suppress_info_messages.IsNull) suppress_info_messages = false;

            if (authority_url.IsNull || string.IsNullOrEmpty(authority_url.Value)) authority_url = DefaultAuthorityUrl;
            if (graph_url.IsNull || string.IsNullOrEmpty(graph_url.Value)) graph_url = DefaultGraphUrl;

            DateTime now = DateTime.UtcNow;
            if (start_time_utc.IsNull) start_time_utc = now.AddMinutes(+300);
            if (end_time_utc.IsNull) end_time_utc = start_time_utc.Value.AddMinutes(+60);

            m_graph_suppress_info_messages = suppress_info_messages.Value;

            string authority = authority_url.Value.TrimEnd('/');
            string graphRoot = graph_url.Value.TrimEnd('/');

            #endregion default values initialization

            #region validations

            currentPhase = "Validating parameters";

            StringBuilder sb_Errors = new StringBuilder();

            if (organizer.IsNull || string.IsNullOrEmpty(organizer.Value))
                sb_Errors.AppendLine("Missing organizer: Please specify @organizer (the M365 mailbox user id or UPN to create the event under)");

            bool hasToken = !access_token.IsNull && !string.IsNullOrEmpty(access_token.Value);
            bool hasClientCreds =
                    !tenant_id.IsNull && !string.IsNullOrEmpty(tenant_id.Value)
                && !client_id.IsNull && !string.IsNullOrEmpty(client_id.Value)
                && !client_secret.IsNull && !string.IsNullOrEmpty(client_secret.Value);

            if (!hasToken && !hasClientCreds)
                sb_Errors.AppendLine("Missing credentials: Please specify either @access_token, or all of @tenant_id, @client_id and @client_secret");

            if (body_format.Value.ToUpper() != "HTML" && body_format.Value.ToUpper() != "TEXT")
                sb_Errors.AppendLine(string.Format("@body_format {0} is invalid. Valid values: TEXT | HTML", body_format.Value));

            GraphMethods graphMethod;
            if (!TryParseEnum(typeof(GraphMethods), method.Value.ToUpper(), out graphMethod))
                sb_Errors.AppendLine(string.Format("@method {0} is invalid. Valid values: {1}", method.Value, String.Join(" | ", Enum.GetNames(typeof(GraphMethods)))));

            if ((graphMethod == GraphMethods.UPDATE || graphMethod == GraphMethods.CANCEL)
                && (event_identifier.IsNull || string.IsNullOrEmpty(event_identifier.Value)))
                sb_Errors.AppendLine(string.Format("@event_identifier is required when @method is '{0}' (it must contain the Graph event id returned when the event was created)", method.Value.ToUpper()));

            if (graphMethod == GraphMethods.REQUEST
                && (recipients.IsNull || string.IsNullOrEmpty(recipients.Value))
                && (optional_recipients.IsNull || string.IsNullOrEmpty(optional_recipients.Value))
                && (resource_recipients.IsNull || string.IsNullOrEmpty(resource_recipients.Value)))
                sb_Errors.AppendLine("Missing recipients: Please specify at least one of @recipients, @optional_recipients or @resource_recipients");

            if (sb_Errors.Length > 0)
                throw new Exception("Unable to send calendar event due to validation error(s): " + sb_Errors);

            #endregion validations

            #region acquire access token

            string bearerToken;
            if (hasToken)
            {
                currentPhase = "Using supplied @access_token";
                bearerToken = access_token.Value;
            }
            else
            {
                currentPhase = "Acquiring Graph access token (client credentials)";
                bearerToken = AcquireGraphToken(authority, tenant_id.Value, client_id.Value, client_secret.Value, graphRoot);
            }

            #endregion acquire access token

            #region build request and call Graph

            // organizer mailbox is URL-path-encoded so UPNs like "user@contoso.com" are safe
            string userSegment = Uri.EscapeDataString(organizer.Value);
            string baseUserUrl = string.Format("{0}/v1.0/users/{1}", graphRoot, userSegment);
            string responseBody;

            switch (graphMethod)
            {
                case GraphMethods.REQUEST:
                {
                    currentPhase = "Constructing event JSON";
                    string payload = BuildEventJson(
                        subject.Value, body, body_format.Value, importance.Value, sensitivity.Value,
                        location, start_time_utc.Value, end_time_utc.Value,
                        recipients, optional_recipients, resource_recipients,
                        use_reminder.Value, reminder_minutes.Value, require_rsvp.Value,
                        create_teams_meeting.Value, all_day_event.Value, true);

                    currentPhase = "Creating event via Graph (POST /events)";
                    responseBody = GraphSend("POST", baseUserUrl + "/events", bearerToken, payload);

                    string newId = ExtractJsonValue(responseBody, "id");
                    if (!string.IsNullOrEmpty(newId)) event_identifier = newId;
                    break;
                }

                case GraphMethods.UPDATE:
                {
                    currentPhase = "Constructing event JSON";
                    string payload = BuildEventJson(
                        subject.Value, body, body_format.Value, importance.Value, sensitivity.Value,
                        location, start_time_utc.Value, end_time_utc.Value,
                        recipients, optional_recipients, resource_recipients,
                        use_reminder.Value, reminder_minutes.Value, require_rsvp.Value,
                        create_teams_meeting.Value, all_day_event.Value, false);

                    string eventSegment = Uri.EscapeDataString(event_identifier.Value);
                    currentPhase = "Updating event via Graph (PATCH /events/{id})";
                    responseBody = GraphSend("PATCH", baseUserUrl + "/events/" + eventSegment, bearerToken, payload);
                    break;
                }

                case GraphMethods.CANCEL:
                {
                    string comment = (cancellation_comment.IsNull) ? "" : cancellation_comment.Value;
                    string payload = "{\"Comment\":" + JsonString(comment) + "}";

                    string eventSegment = Uri.EscapeDataString(event_identifier.Value);
                    currentPhase = "Cancelling event via Graph (POST /events/{id}/cancel)";
                    responseBody = GraphSend("POST", baseUserUrl + "/events/" + eventSegment + "/cancel", bearerToken, payload);
                    break;
                }

                default:
                    throw new Exception("Unhandled method: " + method.Value);
            }

            response_content = string.IsNullOrEmpty(responseBody) ? SqlString.Null : (SqlString)responseBody;

            #endregion build request and call Graph

            if (!suppress_info_messages.Value)
                SqlContext.Pipe.Send(string.Format("Calendar event {0} succeeded. Event Identifier: {1}",
                    method.Value.ToUpper(),
                    event_identifier.IsNull ? "(n/a)" : event_identifier.Value));
        }
        catch (Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            Exception innerEx = ex;
            sb.AppendLine(string.Format("Error while {0}: [{1}] {2}", currentPhase, ex.Source, ex.Message));

            while (ex.InnerException != null)
            {
                ex = ex.InnerException;
                sb.AppendLine(string.Format("[{0}] {1}", ex.Source, ex.Message));
            }

            throw new Exception(sb.ToString(), innerEx);
        }
    }

    #region Graph helper methods

    /// <summary>
    /// Acquires an app-only bearer token from Azure AD using the OAuth2
    /// client-credentials flow with the "/.default" Graph scope.
    /// </summary>
    private static string AcquireGraphToken(string authority, string tenantId, string clientId, string clientSecret, string graphRoot)
    {
        EnsureModernTls();

        string tokenUrl = string.Format("{0}/{1}/oauth2/v2.0/token", authority, Uri.EscapeDataString(tenantId));

        StringBuilder form = new StringBuilder();
        form.Append("client_id=").Append(Uri.EscapeDataString(clientId));
        form.Append("&client_secret=").Append(Uri.EscapeDataString(clientSecret));
        form.Append("&scope=").Append(Uri.EscapeDataString(graphRoot + "/.default"));
        form.Append("&grant_type=client_credentials");

        byte[] formBytes = Encoding.UTF8.GetBytes(form.ToString());

        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(tokenUrl);
        req.Method = "POST";
        req.ContentType = "application/x-www-form-urlencoded";
        req.Accept = "application/json";
        req.ContentLength = formBytes.Length;

        using (Stream reqStream = req.GetRequestStream())
        {
            reqStream.Write(formBytes, 0, formBytes.Length);
        }

        string responseBody = ReadResponse(req);
        string token = ExtractJsonValue(responseBody, "access_token");

        if (string.IsNullOrEmpty(token))
            throw new Exception("Failed to acquire access token: no 'access_token' found in the authority response.");

        return token;
    }

    /// <summary>
    /// Sends an authenticated request to Microsoft Graph and returns the
    /// response body (empty string for 202/204 responses that carry no body).
    /// </summary>
    private static string GraphSend(string httpMethod, string url, string bearerToken, string jsonPayload)
    {
        EnsureModernTls();

        HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
        req.Method = httpMethod;
        req.Accept = "application/json";
        req.Headers.Add("Authorization", "Bearer " + bearerToken);

        if (jsonPayload != null)
        {
            byte[] payloadBytes = Encoding.UTF8.GetBytes(jsonPayload);
            req.ContentType = "application/json; charset=utf-8";
            req.ContentLength = payloadBytes.Length;
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(payloadBytes, 0, payloadBytes.Length);
            }
        }

        return ReadResponse(req);
    }

    /// <summary>
    /// Reads the response body, surfacing the Graph/authority error payload in
    /// the exception message when a non-success status code is returned.
    /// </summary>
    private static string ReadResponse(HttpWebRequest req)
    {
        try
        {
            using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
            {
                return ReadStream(resp);
            }
        }
        catch (WebException wex)
        {
            HttpWebResponse errResp = wex.Response as HttpWebResponse;
            if (errResp != null)
            {
                string errBody = ReadStream(errResp);
                throw new Exception(string.Format("HTTP {0} ({1}) from {2}: {3}",
                    (int)errResp.StatusCode, errResp.StatusDescription, req.RequestUri, errBody), wex);
            }
            throw;
        }
    }

    private static string ReadStream(HttpWebResponse resp)
    {
        Stream s = resp.GetResponseStream();
        if (s == null) return "";
        using (StreamReader reader = new StreamReader(s, Encoding.UTF8))
        {
            return reader.ReadToEnd();
        }
    }

    private static void EnsureModernTls()
    {
        ServicePointManager.SecurityProtocol =
            SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
    }

    /// <summary>
    /// Builds the Graph event JSON body. When <paramref name="includeAttendees"/>
    /// is false only the mutable scalar fields are emitted (used for PATCH so we
    /// don't inadvertently resend the whole attendee list on an update).
    /// </summary>
    private static string BuildEventJson(
        string subject, SqlString body, string bodyFormat, string importance, string sensitivity,
        SqlString location, DateTime startUtc, DateTime endUtc,
        SqlString required, SqlString optional, SqlString resource,
        bool useReminder, int reminderMinutes, bool requireRsvp,
        bool createTeamsMeeting, bool allDayEvent, bool includeAttendees)
    {
        StringBuilder json = new StringBuilder();
        json.Append("{");

        json.Append("\"subject\":").Append(JsonString(subject));

        string contentType = (bodyFormat.ToUpper() == "HTML") ? "HTML" : "Text";
        string bodyContent = body.IsNull ? "" : body.Value;
        json.Append(",\"body\":{\"contentType\":").Append(JsonString(contentType))
            .Append(",\"content\":").Append(JsonString(bodyContent)).Append("}");

        // All-day events must be set to midnight boundaries; Graph requires
        // start/end on date boundaries and the isAllDay flag.
        string startStr = allDayEvent
            ? startUtc.ToString("yyyy-MM-ddT00:00:00")
            : startUtc.ToString("yyyy-MM-ddTHH:mm:ss");
        string endStr = allDayEvent
            ? endUtc.ToString("yyyy-MM-ddT00:00:00")
            : endUtc.ToString("yyyy-MM-ddTHH:mm:ss");

        json.Append(",\"start\":{\"dateTime\":").Append(JsonString(startStr)).Append(",\"timeZone\":\"UTC\"}");
        json.Append(",\"end\":{\"dateTime\":").Append(JsonString(endStr)).Append(",\"timeZone\":\"UTC\"}");
        if (allDayEvent) json.Append(",\"isAllDay\":true");

        json.Append(",\"importance\":").Append(JsonString(MapImportance(importance)));
        json.Append(",\"sensitivity\":").Append(JsonString(MapSensitivity(sensitivity)));

        if (!location.IsNull && !string.IsNullOrEmpty(location.Value))
            json.Append(",\"location\":{\"displayName\":").Append(JsonString(location.Value)).Append("}");

        json.Append(",\"isReminderOn\":").Append(useReminder ? "true" : "false");
        if (useReminder)
            json.Append(",\"reminderMinutesBeforeStart\":").Append(reminderMinutes);

        json.Append(",\"responseRequested\":").Append(requireRsvp ? "true" : "false");

        if (createTeamsMeeting)
            json.Append(",\"isOnlineMeeting\":true,\"onlineMeetingProvider\":\"teamsForBusiness\"");

        if (includeAttendees)
        {
            List<string> attendees = new List<string>();
            AppendAttendees(attendees, required, "required");
            AppendAttendees(attendees, optional, "optional");
            AppendAttendees(attendees, resource, "resource");

            json.Append(",\"attendees\":[").Append(String.Join(",", attendees.ToArray())).Append("]");
        }

        json.Append("}");
        return json.ToString();
    }

    /// <summary>
    /// Parses a semicolon/comma-delimited recipient list where each entry is
    /// either "email" or "Display Name &lt;email&gt;", emitting Graph attendee objects.
    /// </summary>
    private static void AppendAttendees(List<string> target, SqlString list, string attendeeType)
    {
        if (list.IsNull || string.IsNullOrEmpty(list.Value)) return;

        foreach (string rawEntry in list.Value.Split(';', ','))
        {
            string entry = rawEntry.Trim();
            if (entry.Length == 0) continue;

            string address = entry;
            string displayName = null;

            int lt = entry.LastIndexOf('<');
            int gt = entry.LastIndexOf('>');
            if (lt >= 0 && gt > lt)
            {
                address = entry.Substring(lt + 1, gt - lt - 1).Trim();
                displayName = entry.Substring(0, lt).Trim().Trim('"');
            }

            StringBuilder a = new StringBuilder();
            a.Append("{\"type\":").Append(JsonString(attendeeType));
            a.Append(",\"emailAddress\":{\"address\":").Append(JsonString(address));
            if (!string.IsNullOrEmpty(displayName))
                a.Append(",\"name\":").Append(JsonString(displayName));
            a.Append("}}");

            target.Add(a.ToString());
        }
    }

    private static string MapImportance(string importance)
    {
        switch (importance.ToUpper())
        {
            case "LOW": return "low";
            case "HIGH": return "high";
            default: return "normal";
        }
    }

    private static string MapSensitivity(string sensitivity)
    {
        // ICS CLASS values -> Graph sensitivity values
        switch (sensitivity.ToUpper())
        {
            case "PRIVATE": return "private";
            case "CONFIDENTIAL": return "confidential";
            case "PERSONAL": return "personal";
            default: return "normal"; // PUBLIC
        }
    }

    /// <summary>
    /// Serializes a string as a JSON string literal (including the surrounding
    /// quotes), escaping control characters per RFC 8259. Returns "null" (the
    /// JSON literal) for a null input.
    /// </summary>
    private static string JsonString(string value)
    {
        if (value == null) return "null";

        StringBuilder sb = new StringBuilder(value.Length + 2);
        sb.Append('"');
        foreach (char c in value)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (c < ' ')
                        sb.Append("\\u").Append(((int)c).ToString("x4"));
                    else
                        sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    /// <summary>
    /// Extracts the value of the first top-level occurrence of a JSON string
    /// property with the given name. This is a deliberately small extractor
    /// sufficient for pulling "access_token" from the token response and "id"
    /// from the created-event response, avoiding an external JSON dependency.
    /// </summary>
    private static string ExtractJsonValue(string json, string propertyName)
    {
        if (string.IsNullOrEmpty(json)) return null;

        // Match "propertyName" : "value" allowing escaped characters in the value.
        string pattern = "\"" + Regex.Escape(propertyName) + "\"\\s*:\\s*\"((?:\\\\.|[^\"\\\\])*)\"";
        Match m = Regex.Match(json, pattern);
        if (!m.Success) return null;

        return JsonUnescape(m.Groups[1].Value);
    }

    /// <summary>
    /// Decodes the escape sequences inside a captured JSON string literal body
    /// (the text between the surrounding quotes) back into their raw characters.
    /// </summary>
    private static string JsonUnescape(string raw)
    {
        if (raw.IndexOf('\\') < 0) return raw;

        StringBuilder sb = new StringBuilder(raw.Length);
        for (int i = 0; i < raw.Length; i++)
        {
            char c = raw[i];
            if (c != '\\' || i + 1 >= raw.Length)
            {
                sb.Append(c);
                continue;
            }

            char next = raw[++i];
            switch (next)
            {
                case '"': sb.Append('"'); break;
                case '\\': sb.Append('\\'); break;
                case '/': sb.Append('/'); break;
                case 'b': sb.Append('\b'); break;
                case 'f': sb.Append('\f'); break;
                case 'n': sb.Append('\n'); break;
                case 'r': sb.Append('\r'); break;
                case 't': sb.Append('\t'); break;
                case 'u':
                    if (i + 4 < raw.Length)
                    {
                        string hex = raw.Substring(i + 1, 4);
                        int code;
                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber,
                                System.Globalization.CultureInfo.InvariantCulture, out code))
                        {
                            sb.Append((char)code);
                            i += 4;
                        }
                        else
                        {
                            sb.Append(next);
                        }
                    }
                    else
                    {
                        sb.Append(next);
                    }
                    break;
                default:
                    sb.Append(next);
                    break;
            }
        }
        return sb.ToString();
    }

    #endregion Graph helper methods
}
