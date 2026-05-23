# nolendar

Nolendar reads upcoming meetings from Outlook Calendar via the Microsoft Graph API and creates tasks in a Notion database so you have a structured place to capture meeting notes, action items, and related context.

## Features

- Read upcoming meetings from Outlook via Microsoft Graph
- Create linked Notion tasks/pages from meetings
- Create one task/page per meeting
- Use a standardized template for created pages
- Include the agenda from the meeting
- Include backlinks to the calendar event
- Provide a place in the Notion page to capture notes
- Support incremental sync using delta queries
- Handle recurring meetings correctly
- Support multiple calendars
- Set the `Due` property on the Notion page to the meeting date/time
- Support a configurable lookahead window for calendar events:
  - next 24h
  - next 7d
  - default: current day's meetings
- Be invokable from the command line
- Allow the command line to print upcoming meetings for the requested time period
- Work for both personal and work Microsoft accounts
- Store configuration in YAML

### Suggested Notion Properties

Suggested properties for each created Notion task/page:

- Meeting title
- Date/time
- Organizer
- Attendees
- Meeting link
- Status
- Notes
- Action items

### Configurable Behavior

Stored configuration should support:

- Target Notion database
- Mapping of calendar event fields to Notion task fields
- Using a specific Notion page template, if specified
- Meeting filtering to avoid creating junk tasks, such as:
  - Ignore declined meetings
  - Ignore meetings under X minutes
  - Ignore events without attendees
  - Ignore personal calendar events
  - Ignore optional attendance

## Idempotency

Idempotency is a critical engineering concern. Nolendar must not create duplicate Notion pages during sync.

Store the following as properties on the target Notion database:

- Outlook event ID
- Event `changeKey` to track the version of the calendar event

The application should either create these properties on the Notion database or prompt the user to create them.

## APIs

This project uses the following APIs:

- Microsoft Graph API
- Notion API or MCP

## Libraries

- [`msgraph-sdk-typescript`](https://github.com/microsoftgraph/msgraph-sdk-typescript) for Microsoft Graph API access
- [`notion-sdk-js`](https://github.com/makenotion/notion-sdk-js) for Notion API access

## Development

Use TDD (`red -> green -> refactor`). Write tests first, then implement functionality.
