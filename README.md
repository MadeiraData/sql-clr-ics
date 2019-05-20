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

> TBA

## License and copyright

This project is copyrighted by Eitan Blumin, and licensed under the MIT license agreement.

More info in [the license file](https://github.com/EitanBlumin/sql_clr_ics/blob/master/LICENSE).

## Acknowledgements

This project was based mostly on the following stack overflow discussion:

[https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment](https://stackoverflow.com/questions/22734403/send-email-to-outlook-with-ics-meeting-appointment)

