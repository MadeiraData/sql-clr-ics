/*
Post-Deployment Script Template							
--------------------------------------------------------------------------------------
 This file contains SQL statements that will be appended to the build script.		
 Use SQLCMD syntax to include a file in the post-deployment script.			
 Example:      :r .\myfile.sql								
 Use SQLCMD syntax to reference a variable in the post-deployment script.		
 Example:      :setvar TableName MyTable							
               SELECT * FROM [$(TableName)]					
--------------------------------------------------------------------------------------
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
	@subject [nvarchar](4000),
	@body [nvarchar](4000) = null,
	@location [nvarchar](4000) = null,
	@start_time [datetime] = null,
	@end_time [datetime] = null,
	@smtp_server [nvarchar](4000) = null,
	@port [int] = 25,
	@use_ssl [bit] = 0,
	@username [nvarchar](4000) = null,
	@password [nvarchar](4000) = null,
	@use_reminder [bit] = 1,
	@reminder_minutes [int] = 15
WITH EXECUTE AS CALLER
AS
EXTERNAL NAME [sql_clr_ics].[StoredProcedures].[clr_send_ics_invite]
GO