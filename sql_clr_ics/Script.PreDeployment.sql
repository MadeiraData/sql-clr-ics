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
use [master];
GO
IF NOT EXISTS (select * from sys.asymmetric_keys WHERE name = 'sql_clr_ics_pkey')
	create asymmetric key sql_clr_ics_pkey
	from executable file = '$(PathToSignedDLL)'
	--encryption by password = 'vtwjmifewVfnhrYke@ZuhxkumsFT7_&#$!~<avjqn|mnvJhp'
GO
IF NOT EXISTS (select name from sys.syslogins where name = 'sql_clr_ics_login')
	create login sql_clr_ics_login from asymmetric key sql_clr_ics_pkey;
GO
grant unsafe assembly to sql_clr_ics_login;
GO
-- Return execution context to intended target database
USE [$(DatabaseName)];
GO