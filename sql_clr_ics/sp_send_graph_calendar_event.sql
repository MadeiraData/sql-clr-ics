/*
	sql_clr_ics copyright (c) Eitan Blumin
---------------------------------------------------
Source: https://github.com/EitanBlumin/sql_clr_ics
License: MIT (https://github.com/EitanBlumin/sql_clr_ics/blob/master/LICENSE)
---------------------------------------------------
This serves as a wrapper for the CLR stored procedure clr_send_graph_event, with
default values for parameters (which is not possible natively for all data types
in the CLR procedure itself).

Unlike sp_send_calendar_event (which builds an ICS file and sends it over SMTP),
this procedure creates/updates/cancels the calendar event directly in a Microsoft
365 mailbox via the Microsoft Graph API. Graph sends the invitations (and any
cancellation notices) to the attendees on the organizer's behalf.
*/
CREATE PROCEDURE [dbo].[sp_send_graph_calendar_event]
	@tenant_id [nvarchar](255) = NULL,
	@client_id [nvarchar](255) = NULL,
	@client_secret [nvarchar](max) = NULL,
	@access_token [nvarchar](max) = NULL,
	@authority_url [nvarchar](255) = N'https://login.microsoftonline.com',
	@graph_url [nvarchar](255) = N'https://graph.microsoft.com',
	@organizer [nvarchar](255) = NULL,
	@recipients [nvarchar](max) = NULL,
	@optional_recipients [nvarchar](max) = NULL,
	@resource_recipients [nvarchar](max) = NULL,
	@subject [nvarchar](255) = NULL,
	@body [nvarchar](max) = NULL,
	@body_format [nvarchar](20) = N'TEXT',
	@importance [nvarchar](6) = N'Normal',
	@sensitivity [nvarchar](12) = N'Public',
	@location [nvarchar](255) = NULL,
	@start_time_utc [datetime] = NULL,
	@end_time_utc [datetime] = NULL,
	@method [nvarchar](7) = N'REQUEST',
	@use_reminder [bit] = 1,
	@reminder_minutes [int] = 15,
	@require_rsvp [bit] = 1,
	@create_teams_meeting [bit] = 0,
	@all_day_event [bit] = 0,
	@cancellation_comment [nvarchar](max) = NULL,
	@suppress_info_messages [bit] = 0,
	@event_identifier [nvarchar](max) = NULL OUTPUT,
	@response_content [nvarchar](max) = NULL OUTPUT
WITH EXECUTE AS CALLER
AS
SET NOCOUNT ON;
EXEC dbo.[clr_send_graph_event] @tenant_id, @client_id, @client_secret, @access_token, @authority_url, @graph_url, @organizer, @recipients, @optional_recipients, @resource_recipients, @subject, @body, @body_format, @importance, @sensitivity, @location, @start_time_utc, @end_time_utc, @method, @use_reminder, @reminder_minutes, @require_rsvp, @create_teams_meeting, @all_day_event, @cancellation_comment, @suppress_info_messages, @event_identifier OUTPUT, @response_content OUTPUT
