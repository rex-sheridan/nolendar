# nolendar

Nolendar reads upcoming meetings from Outlook Calendar via Microsoft Graph and turns them into structured Notion tasks/pages for meeting notes, action items, and related context.

The project is currently in early development. Milestone 3 is complete: Nolendar can load YAML config, authenticate to Microsoft Graph, validate a target Notion data source, and perform an idempotent initial sync that creates or updates meeting pages without duplicating them.

## Current Status

Implemented now:

- YAML config loading and validation
- CLI commands for config validation, Notion validation, meeting listing, and sync
- Microsoft Graph authentication via:
  - device code flow
  - interactive browser flow using Azure Identity's development application
  - authorization code flow for web-style app registrations
- Notion API integration
- Multi-calendar meeting listing
- Configurable lookahead windows:
  - `today`
  - quantity-based relative ranges like `12h`, `5d`, `2w`, and `3m`
- Normalized meeting output for CLI display
- Notion schema validation for required properties
- Optional creation of missing required Notion properties
- Idempotent initial sync using Outlook event ID and `changeKey`
- Sync filtering for:
  - declined meetings
  - minimum meeting duration
  - meetings without attendees
- Created Notion pages include sections for:
  - Meeting Link
  - Calendar Event
  - Meeting Details
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
- Source URL
- Tags
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
- A Microsoft app registration for Microsoft Graph user authentication
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
  authMode: device_code

notion:
  databaseId: your_notion_database_id
  defaultTags:
    - meeting
    - notes

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
  eventLink: Source URL
  tags: Tags

sync:
  lookahead: today
  statePath: ./.nolendar/state.json
```

### Config Notes

- `microsoft.tenant` supports `common`, `organizations`, or `consumers`
- `microsoft.authMode` supports:
  - `device_code`
  - `interactive_browser`
  - `auth_code`
- `notion.databaseId` is required
- `notion.defaultTags` can be used to apply static tags to each created Notion page
- with the current Notion API model, this value should be the target data source ID used for row queries and page creation
- `calendars` must contain at least one calendar
- `sync.lookahead` defaults to `today`
- `sync.lookahead` also accepts relative ranges with a numeric prefix:
  - `h` for hours
  - `d` for days
  - `w` for weeks
  - `m` for months
  - examples: `12h`, `5d`, `2w`, `3m`
- `sync.statePath` defaults to `.nolendar/state.json`, resolved relative to the config file
- mapping defaults are:
  - `Name`
  - `Due`
  - `Outlook Event ID`
  - `Outlook ChangeKey`
- optional mapping fields:
  - `mapping.eventLink` for a Notion `url` property populated from Outlook `webLink`
  - `mapping.tags` for a Notion `multi_select` property populated from `notion.defaultTags`

### Required Notion Properties

Nolendar validates the following properties on the target Notion data source:

- title property mapped by `mapping.title`
- date property mapped by `mapping.due`
- rich text property mapped by `mapping.eventId`
- rich text property mapped by `mapping.changeKey`
- url property mapped by `mapping.eventLink`, if configured
- multi-select property mapped by `mapping.tags`, if configured and `notion.defaultTags` is non-empty

If these properties are missing, you can either create them yourself or use `--ensure-properties` with `validate-notion` or `sync`.

### Notion Property Mapping

Current Notion page creation supports:

- title from the Outlook meeting subject
- due date from the meeting start/end
- event ID from the Outlook event ID
- change key from the Outlook event `changeKey`
- URL from the Outlook event `webLink` when `mapping.eventLink` is configured
- tags from `notion.defaultTags` when `mapping.tags` is configured
- meeting body content from the full Outlook event body when available, otherwise `bodyPreview`
- Teams join link extraction from the Outlook event body when Graph does not return `onlineMeeting.joinUrl`

Current Notion page body content includes:

- a `Meeting Link` section when a meeting join URL is available
- a `Calendar Event` section when the Outlook `webLink` is available
- a `Meeting Details` section populated from the event body text
- `Notes`
- `Action items`

Example:

```yaml
notion:
  databaseId: your_notion_data_source_id
  defaultTags:
    - meeting
    - sync

mapping:
  title: Name
  due: Due
  eventId: Outlook Event ID
  changeKey: Outlook ChangeKey
  eventLink: Source URL
  tags: Tags
```

## Environment Variables

Microsoft Graph auth may use:

```bash
export MICROSOFT_CLIENT_ID=your_app_registration_client_id
```

For `device_code` and `interactive_browser`, `MICROSOFT_CLIENT_ID` is optional. If omitted, Nolendar falls back to Azure Identity's development application.

For `microsoft.authMode: auth_code`, you also need:

```bash
export MICROSOFT_CLIENT_SECRET=your_client_secret
export MICROSOFT_REDIRECT_URI=http://localhost:8787/auth/callback
```

You can also bypass Nolendar's built-in Microsoft auth flows entirely by supplying a raw bearer token:

```bash
export MICROSOFT_ACCESS_TOKEN=eyJ...
```

If `MICROSOFT_ACCESS_TOKEN` is set, Nolendar uses it directly for Microsoft Graph requests and ignores the configured Microsoft auth mode for that process.

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

- optionally `MICROSOFT_CLIENT_ID`
- optionally `MICROSOFT_CLIENT_SECRET` and `MICROSOFT_REDIRECT_URI` for `auth_code`
- optionally `MICROSOFT_ACCESS_TOKEN` to bypass Nolendar's built-in Microsoft auth flows
- `NOTION_TOKEN` or `NOTION_API_KEY`
- one or more Outlook calendar IDs
- a target Notion data source ID

### Microsoft App Registration

Nolendar supports two Microsoft auth modes.

#### Option 1: `device_code`

Use this if your app registration can enable public client flows.

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

Microsoft config:

```yaml
microsoft:
  tenant: common
  authMode: device_code
```

Environment:

```bash
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

#### Option 2: `auth_code`

Use this if your app registration is configured as a web app and you can create a client secret.

High-level steps:

1. Open the same app registration in Microsoft Entra
2. Go to `Authentication`
3. Under `Platform configurations`, add or confirm a `Web` platform
4. Add a redirect URI such as `http://localhost:8787/auth/callback`
5. Save
6. Go to `Certificates & secrets`
7. Create a new client secret
8. Copy the `Application (client) ID`
9. Copy the secret value immediately after creating it

Microsoft config:

```yaml
microsoft:
  tenant: common
  authMode: auth_code
```

Environment:

```bash
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
export MICROSOFT_CLIENT_SECRET=your_client_secret
export MICROSOFT_REDIRECT_URI=http://localhost:8787/auth/callback
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

Notes:

- `MICROSOFT_REDIRECT_URI` must exactly match a redirect URI registered on the app
- the current implementation requires a localhost HTTP redirect URI for the CLI callback listener
- Nolendar will open a browser for sign-in and wait for the redirect back to the local callback URL

#### Option 3: `interactive_browser`

Use this when you cannot modify the app registration and need a local developer fallback. This mode uses Azure Identity's development application instead of your own Microsoft Entra app registration.

Microsoft config:

```yaml
microsoft:
  tenant: common
  authMode: interactive_browser
```

Environment:

```bash
unset MICROSOFT_CLIENT_ID
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
```

Notes:

- this mode is intended as a developer fallback, not a production setup
- Nolendar will open a browser for sign-in
- because it does not rely on your app registration, it avoids both:
  - `Allow public client flows`
  - client secret creation

You can also use `device_code` without `MICROSOFT_CLIENT_ID` to fall back to Azure Identity's development application, but `interactive_browser` is the better choice when device-code flow is blocked by app-registration policy.

#### Option 4: `MICROSOFT_ACCESS_TOKEN` override

Use this when you already have a short-lived Microsoft Graph bearer token from another client and want Nolendar to use it directly.

Environment:

```bash
export MICROSOFT_ACCESS_TOKEN=eyJ...
```

Notes:

- if `MICROSOFT_ACCESS_TOKEN` is set, it takes precedence over Nolendar's other Microsoft auth modes
- this is useful as a developer fallback when app registration changes are blocked
- the token is short-lived and will need to be refreshed manually
- a common way to get such a token is from Graph Explorer's `Access token` tab after signing in and consenting to calendar permissions
- the token must be valid for Microsoft Graph and include the permissions needed for the request, such as reading calendar events

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
npm run dev -- sync --config nolendar.yml --force-update
```

If you are using `MICROSOFT_ACCESS_TOKEN`, export it before running `list` or `sync`:

```bash
export MICROSOFT_ACCESS_TOKEN=eyJ...
npm run dev -- list --config nolendar.yml --lookahead 5d
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
npm run dev -- list --config nolendar.yml --lookahead 5d
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

Force-update already-synced pages even when the stored Outlook `changeKey` matches:

```bash
npm run dev -- sync --config nolendar.yml --force-update
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

When using `auth_code`, Nolendar opens a browser for Microsoft sign-in and listens for the redirect back to `MICROSOFT_REDIRECT_URI`.

### Current Sync Behavior

The current sync implementation:

- validates the Notion schema before writing
- optionally creates missing required properties with `--ensure-properties`
- filters meetings using:
  - `ignoreDeclined`
  - `minDurationMinutes`
  - `requireAttendees`
- looks up existing Notion pages by the configured Outlook event ID property
- skips updates when the stored `changeKey` matches unless `--force-update` is used
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
- [`@azure/msal-node`](https://www.npmjs.com/package/@azure/msal-node) for Microsoft authorization-code authentication
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
