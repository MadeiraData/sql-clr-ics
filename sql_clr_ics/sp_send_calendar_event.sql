/*
	sql_clr_ics copyright (c) Eitan Blumin
---------------------------------------------------
Source: https://github.com/EitanBlumin/sql_clr_ics
License: MIT (https://github.com/EitanBlumin/sql_clr_ics/blob/master/LICENSE)
---------------------------------------------------
This serves as a wrapper for the CLR stored procedure with default values for parameters
(which is not possible natively for all data types in the CLR procedure itself)
*/
CREATE PROCEDURE [dbo].[sp_send_calendar_event]
	@profile_name [sysname] = NULL,
	@recipients [nvarchar](max) = NULL,
	@copy_recipients [nvarchar](max) = NULL,
	@blind_copy_recipients [nvarchar](max) = NULL,
	@from_address [nvarchar](max) = NULL,
	@reply_to [nvarchar](max) = NULL,
	@subject [nvarchar](255) = N'SQL Server Meeting',
	@body [nvarchar](max) = NULL,
	@body_format [nvarchar](20) = N'TEXT',
	@importance [nvarchar](6) = N'Normal',
	@sensitivity [nvarchar](12) = N'Public',
	@file_attachments [nvarchar](max) = NULL,
	@location [nvarchar](255) = NULL,
	@start_time_utc [datetime] = NULL,
	@end_time_utc [datetime] = NULL,
	@timestamp_utc [datetime] = NULL,
	@method [nvarchar](14) = N'REQUEST',
	@sequence [int] = 0,
	@prod_id [nvarchar](255) = NULL,
	@use_reminder [bit] = 1,
	@reminder_minutes [int] = 15,
	@require_rsvp [bit] = 0,
	@recipients_role [nvarchar](15) = N'REQ-PARTICIPANT',
	@copy_recipients_role [nvarchar](15) = N'OPT-PARTICIPANT',
	@blind_copy_recipients_role [nvarchar](15) = N'NON-PARTICIPANT',
	@smtp_servername [nvarchar](255) = N'localhost',
	@port [int] = 25,
	@enable_ssl [bit] = 0,
	@use_default_credentials [bit] = 0,
	@username [nvarchar](255) = NULL,
	@password [nvarchar](255) = NULL,
	@suppress_info_messages [bit] = 0,
	@event_identifier [nvarchar](255) = NULL OUTPUT
WITH EXECUTE AS CALLER
AS
SET NOCOUNT ON;
EXEC dbo.[clr_send_ics_invite] @profile_name, @recipients, @copy_recipients, @blind_copy_recipients, @from_address, @reply_to, @subject, @body, @body_format, @importance, @sensitivity, @file_attachments, @location, @start_time_utc, @end_time_utc, @timestamp_utc, @method, @sequence, @prod_id, @use_reminder, @reminder_minutes, @require_rsvp, @recipients_role, @copy_recipients_role, @blind_copy_recipients_role, @smtp_servername, @port, @enable_ssl, @use_default_credentials, @username, @password, @suppress_info_messages, @event_identifier OUTPUT
