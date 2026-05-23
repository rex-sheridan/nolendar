# nolendar

Nolendar reads upcoming meetings from Outlook Calendar via Microsoft Graph and turns them into structured Notion tasks/pages for meeting notes, action items, and related context.

The project is currently in early development. Milestone 2 is complete: Nolendar can load YAML config, authenticate to Microsoft Graph using device-code auth, and list upcoming meetings from one or more Outlook calendars on the command line.

## Current Status

Implemented now:

- YAML config loading and validation
- CLI commands for config validation and meeting listing
- Microsoft Graph device-code authentication
- Multi-calendar meeting listing
- Configurable lookahead windows:
  - `today`
  - `24h`
  - `7d`
- Normalized meeting output for CLI display

Planned next:

- Notion database validation
- Idempotent initial sync from Outlook to Notion
- Delta-query incremental sync
- Recurring meeting sync hardening
- Template-based page creation

## Planned Features

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
- Work for both personal and work Microsoft accounts
- Store configuration in YAML

### Suggested Notion Properties

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

## Requirements

- Node.js 20+
- A Microsoft app registration with a client ID for Graph device-code authentication
- Access to the Outlook calendars you want to read

## Installation

```bash
npm install
```

## Configuration

Nolendar reads YAML config. By default, commands look for `nolendar.yml` in the current directory.

Example:

```yaml
microsoft:
  tenant: common

notion:
  databaseId: your_notion_database_id

calendars:
  - id: primary
    name: Primary
  - id: team-calendar-id
    name: Team

filters:
  ignoreDeclined: true
  minDurationMinutes: 15
  requireAttendees: false
  ignorePersonal: false
  ignoreOptionalAttendance: false

mapping:
  title: Name
  due: Due
  eventId: Outlook Event ID
  changeKey: Outlook ChangeKey

sync:
  lookahead: today
  statePath: ./.nolendar/state.json
```

### Config Notes

- `microsoft.tenant` supports `common`, `organizations`, or `consumers`
- `notion.databaseId` is already required, even though Notion sync is not implemented yet
- `calendars` must contain at least one calendar
- `sync.lookahead` defaults to `today`
- `sync.statePath` defaults to `.nolendar/state.json`, resolved relative to the config file
- mapping defaults are:
  - `Name`
  - `Due`
  - `Outlook Event ID`
  - `Outlook ChangeKey`

## Environment Variables

Current Microsoft Graph auth uses device-code flow and requires:

```bash
export MICROSOFT_CLIENT_ID=your_app_registration_client_id
```

Optional:

```bash
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

If `MICROSOFT_GRAPH_SCOPES` is not set, Nolendar uses:

- `Calendars.Read`
- `User.Read`

## Usage

Run the CLI in development mode:

```bash
npm run dev -- validate-config --config nolendar.yml
```

List meetings using the config default lookahead:

```bash
npm run dev -- list --config nolendar.yml
```

Override the lookahead window:

```bash
npm run dev -- list --config nolendar.yml --lookahead 24h
```

Or build the project and run the compiled CLI:

```bash
npm run build
node dist/index.js list --config nolendar.yml --lookahead 7d
```

### Available Commands

- `list`
  - Prints upcoming meetings for the configured calendars and requested time window
- `validate-config`
  - Loads the YAML config and prints the normalized result as JSON

### Example `list` Output

```text
Meetings for today from 2 configured calendar(s) between 2026-05-22T00:00:00.000Z and 2026-05-23T00:00:00.000Z:
2026-05-22T13:00:00.000Z | Planning
  Calendar: Team
  Organizer: Jordan
  Attendees: Riley, morgan@example.com (optional)
  Meeting Link: https://teams.example/join
  Event Link: https://outlook.office.com/calendar/item/...
  Flags: recurring, accepted
```

When using device-code auth, Microsoft will prompt you in the terminal to open `https://microsoft.com/devicelogin` and enter a code.

## Idempotency

Idempotency is a critical engineering concern. Nolendar must not create duplicate Notion pages during sync.

Store the following as properties on the target Notion database:

- Outlook event ID
- Event `changeKey` to track the version of the calendar event

The application should either create these properties on the Notion database or prompt the user to create them.

## APIs

This project is designed around:

- Microsoft Graph API
- Notion API or MCP

## Libraries

- [`msgraph-sdk-typescript`](https://github.com/microsoftgraph/msgraph-sdk-typescript) for Microsoft Graph API access
- [`notion-sdk-js`](https://github.com/makenotion/notion-sdk-js) for Notion API access
- [`@azure/identity`](https://www.npmjs.com/package/@azure/identity) for Microsoft device-code authentication
- [`commander`](https://www.npmjs.com/package/commander) for the CLI
- [`yaml`](https://www.npmjs.com/package/yaml) for config parsing

## Development

Development follows TDD: `red -> green -> refactor`.

Available scripts:

```bash
npm test
npm run test:watch
npm run lint
npm run build
npm run dev -- list --config nolendar.yml
```

## Implementation Plan

The build plan is tracked in [IMPLEMENTATION_PLAN.md](/Users/rex/workspace/nolendar/IMPLEMENTATION_PLAN.md).

Current completed milestones:

- Milestone 1: project scaffold, config loading, basic CLI
- Milestone 2: Microsoft Graph meeting listing
