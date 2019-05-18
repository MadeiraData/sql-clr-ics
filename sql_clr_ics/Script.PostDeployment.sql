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
	@from [nvarchar](4000),
	@to [nvarchar](4000),
	@cc [nvarchar](4000) = null,
	@reply_to [nvarchar](4000) = null,
	@subject [nvarchar](4000),
	@body [nvarchar](4000) = null,
	@location [nvarchar](4000) = null,
	@start_time_utc [datetime] = null,
	@end_time_utc [datetime] = null,
	@timestamp_utc [datetime] = null,
	@smtp_server [nvarchar](4000) = null,
	@port [int] = 25,
	@use_ssl [bit] = 0,
	@username [nvarchar](4000) = null,
	@password [nvarchar](4000) = null,
	@use_reminder [bit] = 1,
	@reminder_minutes [int] = 15,
	@cancel_event_identifier [uniqueidentifier] = null,
	@event_identifier [uniqueidentifier] = null OUTPUT,
	@suppress_info_messages [bit] = 0
WITH EXECUTE AS CALLER
AS
EXTERNAL NAME [sql_clr_ics].[StoredProcedures].[clr_send_ics_invite]
GO