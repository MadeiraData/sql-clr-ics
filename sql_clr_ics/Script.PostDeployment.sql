/*
Post-Deployment Script
---------------------------------------------------
This script re-creates the CLR stored procedure with default values for parameters
(which is not possible natively)
*/
IF OBJECT_ID('clr_send_ics_invite') IS NOT NULL DROP PROCEDURE [dbo].[clr_send_ics_invite]
GO
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
CREATE PROCEDURE [dbo].[clr_send_ics_invite]
	@profile_name [sysname] = NULL,
	@recipients [nvarchar](4000) = NULL,
	@copy_recipients [nvarchar](4000) = NULL,
	@blind_copy_recipients [nvarchar](4000) = NULL,
	@from_address [nvarchar](4000) = NULL,
	@reply_to [nvarchar](4000) = NULL,
	@subject [nvarchar](255) = N'SQL Server Meeting',
	@body [nvarchar](4000) = NULL,
	@body_format [nvarchar](20) = N'TEXT',
	@importance [nvarchar](6) = N'Normal',
	@sensitivity [nvarchar](12) = N'Public',
	@file_attachments [nvarchar](4000) = NULL,
	@location [nvarchar](4000) = NULL,
	@start_time_utc [datetime] = NULL,
	@end_time_utc [datetime] = NULL,
	@timestamp_utc [datetime] = NULL,
	@method [nvarchar](14) = N'REQUEST',
	@sequence [int] = 0,
	@prod_id [nvarchar](4000) = NULL,
	@use_reminder [bit] = 1,
	@reminder_minutes [int] = 15,
	@require_rsvp [bit] = 0,
	@recipients_role [nvarchar](15) = N'REQ-PARTICIPANT',
	@copy_recipients_role [nvarchar](15) = N'OPT-PARTICIPANT',
	@blind_copy_recipients_role [nvarchar](15) = N'NON-PARTICIPANT',
	@smtp_servername [nvarchar](4000) = N'localhost',
	@port [int] = 25,
	@enable_ssl [bit] = 0,
	@use_default_credentials [bit] = 0,
	@username [nvarchar](4000) = NULL,
	@password [nvarchar](4000) = NULL,
	@suppress_info_messages [bit] = 0,
	@event_identifier [nvarchar](4000) = NULL OUTPUT
WITH EXECUTE AS CALLER
AS
EXTERNAL NAME [sql_clr_ics].[StoredProcedures].[clr_send_ics_invite]
GO
