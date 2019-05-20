# sql_clr_ics

Send Calendar Invites from within SQL Server using a CLR stored procedure.

## Prerequisites

The CLR assembly in this project can only be deployed to a database with the **trustworthy** setting on, due to the assembly requiring the **UNSAFE permission set**.

## Installation

If you have SSDT, you can open the SQL server project and publish it to your database of choice.
Alternatively, you can use [this simple installation script](https://github.com/EitanBlumin/sql_clr_ics/blob/master/sql_clr_ics/sql_clr_ics_install.sql) that sets everything up for you in your database of choice.

## Usage

The CLR stored procedure `clr_send_ics_invite` accepts the following parameters:

```
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
	@require_rsvp [bit] = 0,
	@cancel_event_identifier [uniqueidentifier] = null,
	@event_identifier [uniqueidentifier] = null OUTPUT,
	@suppress_info_messages [bit] = 0
```

|Parameter|Type|Default|Description|
|---|---|---|---|
| `@from` | nvarchar(4000) | _no default_ | Must be a valid single e-mail address from which the invite will be sent. |
| `@to` | nvarchar(4000) | _no default_ | Accepts a list of e-mail addresses (at least one) to be invited as required partisipants, separated by either a comma or a semicolon. |
| `@cc` | nvarchar(400) | _null_ | Optional parameter. Accepts a list of e-mail addresses (at least one) to be used as CC, separated by either a comma or a semicolon. |
| `@reply_to` | nvarchar(4000) | _null_ | Optional parameter. Accepts an e-mail address to be used as the Reply To address (if different from the `@from` address. |
| `@subject` | nvarchar(4000) | _no default_ | Mandatory parameter. A text string to be used as the meeting / e-mail's subject. |
| `@body` | nvarchar(4000) | _null_ | Optional parameter. A text string to be used as the e-mail's HTML body. |
| `@location` | nvarchar(4000) | _null_ | Optional parameter. Sets the location for the meeting. |
| `@start_time_utc` | datetime | _UTC now + 5 hours_ | Optional parameter. Sets the start time (in UTC) of the meeting. If not specified, by default will be set as **UTC now + 5 hours**. |
| `@end_time_utc` | datetime | _@start_time_utc + 1 hour_ | Optional parameter. Sets the end time (in UTC) of the meeting. If not specified, by default will be set as **`@start_time_utc` + 1 hour**. |
| `@timestamp_utc` | datetime | _UTC now_ | Optional parameter. Sets the DTSTAMP section of the iCal (usually used for consistent updating of meeting invites). If not specified, by default will be set as **UTC now**. |
| `@smtp_server` | nvarchar(4000) | _localhost_ | Optional parameter. Sets the SMTP host name to be used for sending the e-mail. If not specified, by default will be set as **"localhost"**. |
| `@port` | int | _25_ | Optional parameter. Sets the SMTP port to be used for sending the e-mail. If not specified, by default will be set as **25**. |
| `@use_ssl` | bit | _0_ | Optional parameter. Sets whether to use SSL authentication for the SMTP server. If not specified, by default will be set as **0 (false)**. |
| `@username` | nvarchar(4000) | _null (use current Network Credentials)_ | Optional parameter. Sets the username to use when authenticating against the SMTP server. If not specified, by default the **current Network Credentials** will be used (of the SQL Server service). |
| `@password` | nvarchar(4000) | _empty password_ | Optional parameter. Sets the password to use when authenticating against the SMTP server. Only used when `@username` is also specified. By default, will use **empty password**. |
| `@use_reminder` | bit | _1_ | Optional parameter. Sets whether to set a reminder for the meeting. By default is set to **1 (true)**. |
| `@reminder_minutes` | int | _15_ | If `@use_reminder` is enabled, this parameter will be used for setting the reminder time in minutes. By default is set to **15**. |
| `@require_rsvp` | bit | _0_ | If set to 0 (false), then participants will not be required to respond with RSVP, and their participation is automatically set as ACCEPTED. If set to 1 (true), then participants will be required to respond with RSVP, and their participation is automatically set as NEEDS-ACTION. By default set to **0 (false)**. |
| `@cancel_event_identifier` | uniqueidentifier | _null_ | You may specify a value for this parameter, if you want to cancel an event that you've already sent. Use the corresponding event's identifier. |
| `@event_identifier` | uniqueidentifier | _null_ | Output parameter. Returns the event's GUID, which can later be used for cancellation. If `@cancel_event_identifier` was specified, the same GUID will be returned. |
| `@suppress_info_messages` | bit | _0_ | If set to 0, an informational message will be printed upon successful delivery of the invitation ( ex. "Mail Sent. Event Identifier: 1234-1234-1234-1234" ). If set to 1, this message will not be printed. By default is set to **0 (false)**. |

## License and copyright

This project is copyrighted by Eitan Blumin, and licensed under the MIT license agreement.

More info in [the license file](https://github.com/EitanBlumin/sql_clr_ics/blob/master/LICENSE).

## Acknowledgements

This project was based mostly on the following stack overflow discussion:

[https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment](https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment)

Also used the iCal specification for further improvements:

[https://www.kanzaki.com/docs/ical/](https://www.kanzaki.com/docs/ical/)
