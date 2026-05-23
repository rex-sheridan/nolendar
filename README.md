# nolendar

Nolendar reads upcoming meetings from Outlook Calendar via Microsoft Graph and turns them into structured Notion tasks/pages for meeting notes, action items, and related context.

The project is currently in early development. Milestone 3 is complete: Nolendar can load YAML config, authenticate to Microsoft Graph, validate a target Notion data source, and perform an idempotent initial sync that creates or updates meeting pages without duplicating them.

## Current Status

Implemented now:

- YAML config loading and validation
- CLI commands for config validation, Notion validation, meeting listing, and sync
- Microsoft Graph device-code authentication
- Notion API integration
- Multi-calendar meeting listing
- Configurable lookahead windows:
  - `today`
  - `24h`
  - `7d`
- Normalized meeting output for CLI display
- Notion schema validation for required properties
- Optional creation of missing required Notion properties
- Idempotent initial sync using Outlook event ID and `changeKey`
- Sync filtering for:
  - declined meetings
  - minimum meeting duration
  - meetings without attendees
- Created Notion pages include sections for:
  - Agenda
  - Notes
  - Action items

Planned next:

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
- A Notion integration token with access to the target data source

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
- `notion.databaseId` is required
- with the current Notion API model, this value should be the target data source ID used for row queries and page creation
- `calendars` must contain at least one calendar
- `sync.lookahead` defaults to `today`
- `sync.statePath` defaults to `.nolendar/state.json`, resolved relative to the config file
- mapping defaults are:
  - `Name`
  - `Due`
  - `Outlook Event ID`
  - `Outlook ChangeKey`

### Required Notion Properties

Nolendar validates the following properties on the target Notion data source:

- title property mapped by `mapping.title`
- date property mapped by `mapping.due`
- rich text property mapped by `mapping.eventId`
- rich text property mapped by `mapping.changeKey`

If these properties are missing, you can either create them yourself or use `--ensure-properties` with `validate-notion` or `sync`.

## Environment Variables

Current Microsoft Graph auth uses device-code flow and requires:

```bash
export MICROSOFT_CLIENT_ID=your_app_registration_client_id
```

Notion API access requires one of:

```bash
export NOTION_TOKEN=secret_xxx
```

or

```bash
export NOTION_API_KEY=secret_xxx
```

Optional:

```bash
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

If `MICROSOFT_GRAPH_SCOPES` is not set, Nolendar uses:

- `Calendars.Read`
- `User.Read`

## Credential Setup

To run a live end-to-end sync, you need:

- `MICROSOFT_CLIENT_ID`
- `NOTION_TOKEN` or `NOTION_API_KEY`
- one or more Outlook calendar IDs
- a target Notion data source ID

### Microsoft App Registration

Nolendar uses Microsoft Graph device-code authentication, so the app registration should be configured as a public client application.

High-level steps:

1. Open the Microsoft Entra admin center at `https://entra.microsoft.com/`
2. Go to `Microsoft Entra ID` -> `Identity` -> `Applications` -> `App registrations`
3. Select `New registration`
4. Enter an app name such as `nolendar`
5. Choose the supported account type:
   - use `Accounts in any organizational directory and personal Microsoft accounts` if you want both work and personal accounts
   - use a narrower option if you only need one tenant or only work accounts
6. Leave `Redirect URI` empty
7. Create the app registration
8. Open `Authentication`
9. Under `Advanced settings`, set `Allow public client flows` to `Yes`
10. Copy the `Application (client) ID` and export it as `MICROSOFT_CLIENT_ID`

Notes:

- device-code flow does not require a client secret
- Nolendar currently uses delegated Graph scopes for user sign-in
- the device-code prompt will direct you to `https://microsoft.com/devicelogin`

### Microsoft Environment

```bash
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

### Notion Integration Token

1. Create a Notion integration in the Notion developer dashboard
2. Copy the integration token
3. Share the target Notion database/data source with that integration
4. Export the token:

```bash
export NOTION_TOKEN=secret_xxx
```

### Outlook Calendar IDs

For a first live run, you can usually start with:

```yaml
calendars:
  - id: primary
    name: Primary
```

If you want a non-primary calendar, you need its Graph calendar ID. Nolendar does not yet have a `list-calendars` command, so today that ID must come from Graph Explorer or another Graph client.

### Notion Data Source ID

The current implementation expects `notion.databaseId` to be the target Notion data source ID used for querying and page creation.

If you only have the parent Notion database/container ID, you may still need the underlying data source ID. In the Notion UI, open the database settings and use the data source management UI to copy the data source ID.

### Live Run Checklist

Once the values are set, the usual sequence is:

```bash
npm run dev -- init
npm run dev -- validate-config --config nolendar.yml
npm run dev -- validate-notion --config nolendar.yml
npm run dev -- sync --config nolendar.yml --dry-run
npm run dev -- sync --config nolendar.yml --ensure-properties
```

## Usage

Run the CLI in development mode:

```bash
npm run dev -- init
npm run dev -- validate-config --config nolendar.yml
```

Generate a starter config at the default location:

```bash
npm run dev -- init
```

Generate a starter config at a custom path:

```bash
npm run dev -- init --config ./config/nolendar.yml
```

Overwrite an existing config file:

```bash
npm run dev -- init --force
```

Validate the generated config:

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

Validate the target Notion data source:

```bash
npm run dev -- validate-notion --config nolendar.yml
```

Create missing required Notion properties when possible:

```bash
npm run dev -- validate-notion --config nolendar.yml --ensure-properties
```

Preview sync actions without changing Notion:

```bash
npm run dev -- sync --config nolendar.yml --dry-run
```

Run the initial sync and auto-create missing required properties:

```bash
npm run dev -- sync --config nolendar.yml --ensure-properties
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
- `validate-notion`
  - Validates access to the configured Notion data source and checks required properties
- `sync`
  - Lists meetings for the requested window and creates, updates, or skips Notion pages based on idempotency checks
- `init`
  - Writes a starter `nolendar.yml` config file

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

### Current Sync Behavior

The current sync implementation:

- validates the Notion schema before writing
- optionally creates missing required properties with `--ensure-properties`
- filters meetings using:
  - `ignoreDeclined`
  - `minDurationMinutes`
  - `requireAttendees`
- looks up existing Notion pages by the configured Outlook event ID property
- skips updates when the stored `changeKey` matches
- updates existing pages when the `changeKey` differs
- creates new pages when no matching event ID is found

### Current Sync Limitations

Not implemented yet:

- delta-query incremental sync
- recurring meeting edge-case hardening
- template-page cloning
- filters for `ignorePersonal` and `ignoreOptionalAttendance`
- MCP-based Notion integration

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
npm run dev -- validate-notion --config nolendar.yml
npm run dev -- sync --config nolendar.yml --dry-run
```

## Implementation Plan

The build plan is tracked in [IMPLEMENTATION_PLAN.md](/Users/rex/workspace/nolendar/IMPLEMENTATION_PLAN.md).

Current completed milestones:

- Milestone 1: project scaffold, config loading, basic CLI
- Milestone 2: Microsoft Graph meeting listing
- Milestone 3: Notion schema validation and idempotent initial sync
