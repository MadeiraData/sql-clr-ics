/*
 Pre-Deployment Script Template							
--------------------------------------------------------------------------------------
 This file contains SQL statements that will be executed before the build script.	
 Use SQLCMD syntax to include a file in the pre-deployment script.			
 Example:      :r .\myfile.sql								
 Use SQLCMD syntax to reference a variable in the pre-deployment script.		
 Example:      :setvar TableName MyTable							
               SELECT * FROM [$(TableName)]					
--------------------------------------------------------------------------------------
*/
-- Make sure clr is enabled
DECLARE @InitAdvanced INT

IF EXISTS (select * from sys.configurations where name IN ('clr enabled') and value_in_use = 0)
BEGIN
	SELECT @InitAdvanced = CONVERT(int, value) FROM sys.configurations WHERE name = 'show advanced options';

	IF @InitAdvanced = 0
	BEGIN
		EXEC sp_configure 'show advanced options', 1;
		RECONFIGURE;
	END

	IF EXISTS (select * from sys.configurations where name = 'clr enabled' and value_in_use = 0)
	BEGIN
		EXEC sp_configure 'clr enabled', 1;
		RECONFIGURE;
	END

	IF @InitAdvanced = 0
	BEGIN
		EXEC sp_configure 'show advanced options', 0;
		RECONFIGURE;
	END
END
GO
use [master];
GO
-- Create symmetric key from DLL
IF NOT EXISTS (
		select * from master.sys.asymmetric_keys 
		WHERE 
			name = 'sql_clr_ics_pkey'
			--thumbprint = 0xC5022B1D1415FC7A
		)
	create asymmetric key sql_clr_ics_pkey
	from executable file = '$(PathToSignedDLL)'
	--encryption by password = 'vtwjmifewVfnhrYke@ZuhxkumsFT7_&#$!~<avjqn|mnvJhp'
GO
-- Create server login from symmetric key
IF NOT EXISTS (select name from master.sys.syslogins where name = 'sql_clr_ics_login')
	create login sql_clr_ics_login from asymmetric key sql_clr_ics_pkey;
GO
-- Grant UNSAFE ASSEMBLY permissions to login which was created from DLL signing key
grant unsafe assembly to sql_clr_ics_login;
GO
-- Return execution context to intended target database
USE [$(DatabaseName)];
GO