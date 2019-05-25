# sql_clr_ics: clr_send_ics_invite, sp_send_calendar_event

Send Calendar Event / Appointment Invitations (iCal formatted file) from within SQL Server using a CLR stored procedure

## Prerequisites

The CLR assembly in this project can only be deployed to a database with the **trustworthy** setting on, due to the assembly requiring the **UNSAFE permission set**.

## Installation

If you have SSDT, you can open the SQL server project and publish it to your database of choice.
Alternatively, you can use [this simple installation script](https://github.com/EitanBlumin/sql_clr_ics/blob/master/sql_clr_ics/sql_clr_ics_install.sql) that sets everything up for you in your database of choice.

## Syntax

```
exec sp_send_calendar_event
	[   [ @profile_name = ] 'profile_name' ]
	[ , [ @recipients = ] 'recipients [ ; ...n ]' ]
	[ , [ @copy_recipients = ] 'copy_recipients [ ; ...n ]' ]
	[ , [ @blind_copy_recipients = ] 'blind_copy_recipients [ ; ...n ]' ]
	[ , [ @from_address = ] 'from_address' ]
	[ , [ @reply_to = ] 'reply_to' ]
	[ , [ @subject = ] 'subject' ]
	[ , [ @body = ] 'body' ]
	[ , [ @body_format = ] 'body_format' ]
	[ , [ @importance = ] 'importance' ]
	[ , [ @sensitivity = ] 'sensitivity' ]
	[ , [ @file_attachments = ] 'file_attachments [ ; ...n ]' ]
	[ , [ @location = ] 'location' ]
	[ , [ @start_time_utc = ] 'start_time_utc' ]
	[ , [ @end_time_utc = ] 'end_time_utc' ]
	[ , [ @timestamp_utc = ] 'timestamp_utc' ]
	[ , [ @method = ] 'method' ]
	[ , [ @sequence = ] sequence ]
	[ , [ @prod_id = ] 'prod_id' ]
	[ , [ @use_reminder = ] use_reminder ]
	[ , [ @reminder_minutes = ] reminder_minutes ]
	[ , [ @require_rsvp = ] require_rsvp ]
	[ , [ @recipients_role = ] 'recipients_role' ]
	[ , [ @copy_recipients_role = ] 'copy_recipients_role' ]
	[ , [ @blind_copy_recipients_role = ] 'blind_copy_recipients_role' ]
	[ , [ @smtp_servername = ] 'smtp_servername' ]
	[ , [ @port = ] port ]
	[ , [ @enable_ssl = ] enable_ssl ]
        [ , [ @use_default_credentials = ] use_default_credentials ]
	[ , [ @username = ] username ]
	[ , [ @password = ] password ]
	[ , [ @suppress_info_messages = ] suppress_info_messages ]
	[ , [ @event_identifier = ] event_identifier [ OUTPUT ] ]
```

## Arguments  

`[ @profile_name = ] 'profile_name'`

 Is the name of the profile to send the message from. The *profile_name* is of type **sysname**, with a default of NULL. The *profile_name* must be the name of an existing Database Mail profile. When no *profile_name* is specified, **clr_send_ics_invite** checks whether **@from_address** was specified. If not, it uses the default public profile for the **msdb** database. If **@from_address** wasn't specified, and there is no default public profile for the database, **@profile_name** must be specified.  
  
`[ @recipients = ] 'recipients'`

 Is a semicolon-delimited list of e-mail addresses to send the message to. The recipients list is of type **nvarchar(max)**. Although this parameter is optional, at least one of **@recipients**, **@copy_recipients**, or **@blind_copy_recipients** must be specified, or **clr_send_ics_invite** returns an error.  
  
`[ @copy_recipients = ] 'copy_recipients'`

 Is a semicolon-delimited list of e-mail addresses to carbon copy the message to. The copy recipients list is of type **nvarchar(max)**. Although this parameter is optional, at least one of **@recipients**, **@copy_recipients**, or **@blind_copy_recipients** must be specified, or **clr_send_ics_invite** returns an error.  
  
`[ @blind_copy_recipients = ] 'blind_copy_recipients'`

 Is a semicolon-delimited list of e-mail addresses to blind carbon copy the message to. The blind copy recipients list is of type **nvarchar(max)**. Although this parameter is optional, at least one of **@recipients**, **@copy_recipients**, or **@blind_copy_recipients** must be specified, or **clr_send_ics_invite** returns an error.  
  
`[ @from_address = ] 'from_address'`

 Is the value of the 'from address' of the email message, and the organizer of the calendar meeting. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **nvarchar(max)**. If no parameter is specified, the default is NULL.
  
`[ @reply_to = ] 'reply_to'`

 Is the value of the 'reply to address' of the email message. It accepts only one email address as a valid value. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **nvarchar(max)**. If no parameter is specified, the default is NULL.  
  
`[ @subject = ] 'subject'`

 Is the subject of the e-mail message. The subject is of type **nvarchar(255)**. If no subject is specified, the default is 'SQL Server Meeting'.  
  
`[ @body = ] 'body'`

 Is the body of the e-mail message. The message body is of type **nvarchar(max)**, with a default of NULL.  
  
`[ @body_format = ] 'body_format'`

 Is the format of the message body. The parameter is of type **varchar(20)**, with a default of NULL. When specified, the headers of the outgoing message are set to indicate that the message body has the specified format. The parameter may contain one of the following values:  
  
-   TEXT
-   HTML  
  
 Defaults to TEXT.  
  
`[ @importance = ] 'importance'`

 Is the importance of the message. The parameter is of type **varchar(6)**. The parameter may contain one of the following values:  
  
-   Low 
-   Normal  
-   High  
  
 Defaults to Normal.  
  
`[ @sensitivity = ] 'sensitivity'`

 Is the sensitivity classification of the message. The parameter is of type **nvarchar(12)**. The parameter may contain one of the following values:  
  
-   Public
-   Private    
-   Confidential  
  
 Defaults to Public.  
  
`[ @file_attachments = ] 'file_attachments'`

 Is a semicolon-delimited list of file names to attach to the e-mail message. Files in the list must be specified as absolute paths. The attachments list is of type **nvarchar(max)**. By default, Database Mail limits file attachments to 1 MB per file.  

`[ @location = ] 'location'`

 Is the location of the calendar meeting. The parameter is of type **nvarchar(255)**, with a default of NULL.
	
`[ @start_time_utc = ] 'start_time_utc'`

 Is the start time of the calendar meeting, in UTC. The parameter is of type **datetime**. If the parameter is not specified, it defaults to **@timestamp_utc** + 5 hours.
 
`[ @end_time_utc = ] 'end_time_utc'`

 Is the end time of the calendar meeting, in UTC. The parameter is of type **datetime**. If the parameter is not specified, it defaults to **@start_time_utc** + 1 hour.
 
`[ @timestamp_utc = ] 'timestamp_utc'`

 Is the DTSTAMP property of the calendar meeting, in UTC. The parameter is of type **datetime**. If the parameter is not specified, it defaults to current UTC time.

`[ @method = ] 'method'`

 Is the method of the calendar event message. The parameter is of type **nvarchar(14)**. The parameter may contain one of the following values:  
  
-   PUBLISH
-   REQUEST
-   REPLY
-   CANCEL
-   ADD
-   REFRESH
-   COUNTER
-   DECLINECOUNTER

 Defaults to REQUEST.  

`[ @sequence = ] sequence`

 Is the sequence of the calendar event message. The parameter is of type **int**, with a default of 0. Unless **@method** is specified as 'CANCEL', in which case the default would be 1. Proper usage of this parameter is important when updating existing calendar events, since each consecutive update must have a higher sequence number than the one before it.
 
`[ @prod_id = ] 'prod_id'`

 Is the PRODID property of the calendar meeting. The parameter is of type **nvarchar(255)**, with a default of 'Schedule a Meeting'.
 
`[ @use_reminder = ] use_reminder`

 Determines whether to add a reminder to the event. The parameter is of type **bit**, with a default of 1 (true).
 
`[ @reminder_minutes = ] reminder_minutes`

 Is the number of minutes to set for the event reminder. The parameter is of type **int**, with a default of 15.
 
`[ @require_rsvp = ] require_rsvp`

 Determines whether participants are required to respond with an RSVP. The parameter is of type **bit**, with a default of 0 (false). If this parameter equals to 0 (false), then all participants are assumed to have accepted their invitation, without requesting a response.
 
`[ @recipients_role = ] 'recipients_role'`

 Is the meeting role for the participants specified in the **@recipients** parameter. The parameter is of type **nvarchar(15)**. The parameter may contain one of the following values:
 
- REQ-PARTICIPANT
- OPT-PARTICIPANT
- NON-PARTICIPANT
- CHAIR

Defaults to REQ-PARTICIPANT.
 
`[ @copy_recipients_role = ] 'copy_recipients_role'`

 Is the meeting role for the participants specified in the **@copy_recipients** parameter. The parameter is of type **nvarchar(15)**. The parameter may contain one of the following values:
 
- REQ-PARTICIPANT
- OPT-PARTICIPANT
- NON-PARTICIPANT
- CHAIR

Defaults to OPT-PARTICIPANT.

`[ @blind_copy_recipients_role = ] 'blind_copy_recipients_role'`

 Is the meeting role for the participants specified in the **@blind_copy_recipients** parameter. The parameter is of type **nvarchar(15)**. The parameter may contain one of the following values:
 
- REQ-PARTICIPANT
- OPT-PARTICIPANT
- NON-PARTICIPANT
- CHAIR

Defaults to NON-PARTICIPANT.

`[ @smtp_servername = ] 'smtp_servername'`

 Is the SMTP server name to be used for sending the e-mail message. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **nvarchar(255)**. If no parameter is specified, and no mail profile was used, the default is 'localhost'.
 
`[ @port = ] port`

 Is the SMTP server port to be used for sending the e-mail message. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **int**. If no parameter is specified, and no mail profile was used, the default is 25.

`[ @enable_ssl = ] enable_ssl`

 Determines whether the SMTP server should use SSL authentication. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **bit**. If no parameter is specified, and no mail profile was used, the default is 0 (false).

`[ @use_default_credentials = ] use_default_credentials`

 Determines whether the SMTP server should use its default network credentials. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **bit**. If no parameter is specified, and no mail profile was used, the default is 0 (false). If **@username** is specified, this parameter is ignored.
 
`[ @username = ] username`

 Is the userame to be used when authenticating with the SMTP server. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **nvarchar(255)**. If no parameter is specified, and no mail profile was used, the default is to use the server's default network credentials instead.
 
`[ @password = ] password`

 Is the password to be used when authenticating with the SMTP server. This is an optional parameter used to override the settings in the mail profile (or if no mail profile was specified). This parameter is of type **nvarchar(255)**. If no parameter is specified, the default is to use an empty string for the password.

| **NOTE:**  Unfortunately, since MSDB doesn't allow access to the mail profile passwords, it's impossible to utilize an existing mail profile for getting the password for an SMTP server. Therefore, unless you want to use an empty password or default network credentials, *you must specify a value for this parameter*. |
| --- |

`[ @suppress_info_messages = ] suppress_info_messages`

 Determines whether to *NOT* display the success message after sending the e-mail. The parameter is of type **bit**, with a default of 0 (false).
	
`[ @event_identifier = ] event_identifier [ OUTPUT ]`

 Optional output parameter returns the *event_identifier* of the calendar meeting. You may also override this value by specifying a parameter with a non-null value for it, in order to uniquely identify a calendar event. If no *event_identifier* was specified, a Global Unique Identifier (Guid) will automatically be generated instead. This parameter must be specified when **@method** is 'CANCEL'. The *event_identifier* is of type **nvarchar(255)**.
  
## Result Sets  

 On success, returns the message "Mail Sent. Event Identifier: %s" (where %s is replaced with the sent **@event_identifier**), unless **@suppress_info_messages** is specified as 1 (true).
 
 On Failure, returns an error message specifying the problem.

## Remarks

I did my best to align the parameters of this procedure with Microsoft's **sp_send_dbmail** procedure. Unfortunately, since this is a CLR procedure, there are limitations to what can be done. Specifically, it's impossible to define default values for parameters of type **nvarchar(max)** and **varchar(max)**, and so I had to create a wrapper procedure in T-SQL instead.

Even though I tried to utilize Microsoft's Database Mail Profile mechanics, I couldn't get access to the account passwords (which is probably a good thing), and so the **@password** parameter becomes mandatory (unless you want to use an empty password or the server's default network credentials).

I also didn't implement any functionality involving multiple accounts per profile to be used as "failover" accounts. So only the first account per profile is used.

## Examples

### A. Send a calendar invitation with RSVP requirement

```
DECLARE @EventID nvarchar(255)
 
EXEC clr_send_ics_invite
        @from_address = N'the_organizer@gmail.com',
        @recipients = N'someone@gmail.com,otherguy@outlook.com',
        @subject = N'let us meet for pizza!',
        @body = N'<h1>Pizza!</h1><p>Bring your own beer!</p>',
        @body_format = N'HTML',
        @location = N'The Pizza place at Hank and Errison corner',
        @start_time_utc = '2019-07-02 19:00',
        @end_time_utc = '2019-07-02 23:00',
        @timestamp_utc = '2019-03-30 18:00',
        @smtp_server = 'smtp.gmail.com',
        @port = 465,
        @enable_ssl = 1,
        @username = N'the_organizer@gmail.com',
        @password = N'NotActuallyMyPassword',
        @use_reminder = 1,
        @reminder_minutes = 30,
        @require_rsvp = 1,
        @event_identifier = @EventID OUTPUT
 
SELECT EventID = @EventID
```

### B. Cancel the previously sent invitation

```
EXEC clr_send_ics_invite
        @from_address = N'the_organizer@gmail.com',
        @recipients = N'someone@gmail.com,otherguy@outlook.com',
        @subject = N'let us meet for pizza!',
        @body = N'<h1>Pizza!</h1><p>Bring your own beer!</p>',
        @body_format = N'HTML',
        @location = N'The Pizza place at Hank and Errison corner',
        @start_time_utc = '2019-07-02 19:00',
        @end_time_utc = '2019-07-02 23:00',
        @timestamp_utc = '2019-03-30 18:00',
        @smtp_server = 'smtp.gmail.com',
        @port = 465,
        @enable_ssl = 1,
        @username = N'the_organizer@gmail.com',
        @password = N'NotActuallyMyPassword',
        @require_rsvp = 1,
        @cancel_event_identifier = @EventID,
        @event_identifier = @EventID OUTPUT
 
SELECT EventID = @EventID
```

### C. Send an automated calendar invitation without RSVP requirement (i.e. participants are auto-accepted)

```
DECLARE @EventID nvarchar(255)
 
EXEC clr_send_ics_invite
        @from_address = N'sla_bot@company.com',
        @recipients = N'employee1@company.com,employee2@company.com',
        @subject = N'Weekly SLA Shift',
        @body = N'<h1>You are on-call this week!</h1><p>This is an automated message</p>',
        @body_format = N'HTML',
        @location = N'Our offices',
        @start_time_utc = '2019-07-01 00:00',
        @end_time_utc = '2019-07-04 23:59',
        @timestamp_utc = '2019-05-01 00:00',
        @smtp_server = 'smtp.company.com',
        @port = 587,
        @enable_ssl = 1,
        @username = N'sla_bot@company.com',
        @password = N'SomethingPassword',
        @use_reminder = 1,
        @reminder_minutes = 300,
        @require_rsvp = 0,
        @event_identifier = @EventID OUTPUT
 
SELECT EventID = @EventID
```
 
## License and copyright

This project is copyrighted by Eitan Blumin, and licensed under the MIT license agreement.

More info in [the license file](https://github.com/EitanBlumin/sql_clr_ics/blob/master/LICENSE).

## Acknowledgements

This project was based mostly on the following stack overflow discussion:

[https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment](https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment)

Also used the iCal specification for further improvements:

[https://www.kanzaki.com/docs/ical/](https://www.kanzaki.com/docs/ical/)

## See Also  
 [sp_send_dbmail](https://docs.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-send-dbmail-transact-sql)   
 [clr_http_request](https://github.com/EitanBlumin/ClrHttpRequest)   
 [clr_wmi_request](https://github.com/EitanBlumin/ClrWmiRequest)
