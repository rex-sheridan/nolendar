# nolendar

Nolendar reads upcoming meetings from Outlook Calendar via Microsoft Graph and turns them into structured Notion pages for notes, action items, and related meeting context.

## Current Status

Implemented now:

- Outlook meeting listing from one or more Microsoft calendars
- Notion page creation and update with idempotency based on Outlook event ID and `changeKey`
- Incremental sync using Microsoft Graph `calendarView/delta`
- Deleted-event archiving and recurring-meeting handling
- Notion template support via:
  - `notion.templatePageId`
  - `notion.dataSourceTemplate`
- Participant relation syncing through a separate Notion People data source
- Configurable filters, tags, assignee resolution, page icons, and page content sections
- CLI timings for Microsoft Graph and Notion API calls

Remaining work is mostly polish:

- auth/runtime cleanup
- helper commands like calendar discovery
- optional MCP-based Notion integration

## Quickstart

1. Install dependencies:

```bash
npm install
```

2. Generate a starter config:

```bash
npm run dev -- init
```

3. Follow [AUTHENTICATION.md](AUTHENTICATION.md) to set up:
   - Microsoft authentication
   - your Notion token
   - Outlook calendar IDs
   - the target Notion data source ID

4. Validate config:

```bash
npm run dev -- validate-config --config nolendar.yml
```

5. Validate Notion access and schema:

```bash
npm run dev -- validate-notion --config nolendar.yml
```

6. Preview sync without writing:

```bash
npm run dev -- sync --config nolendar.yml --dry-run
```

7. Run sync:

```bash
npm run dev -- sync --config nolendar.yml
```

## Requirements

- Node.js 20+
- access to the Outlook calendars you want to read
- a Notion token with access to the target data source

Detailed auth and credential setup lives in [AUTHENTICATION.md](/Users/rex/workspace/nolendar/AUTHENTICATION.md).

## Configuration

Nolendar reads YAML config. By default, commands look for `nolendar.yml` in the current directory.

Example:

```yaml
microsoft:
  tenant: common
  authMode: device_code

notion:
  databaseId: your_notion_data_source_id
  # Choose at most one template mode:
  # templatePageId: your_template_page_id
  # dataSourceTemplate:
  #   type: default
  #   # templateId: your_data_source_template_id
  #   # timezone: America/New_York
  peopleDataSource:
    databaseId: your_people_data_source_id
    nameProperty: Name
    emailProperty: Email Address
  defaultTags:
    - meeting
  defaultAssigneeEmail: you@example.com
  pageIcon:
    type: emoji
    emoji: "📝"
  pageContent:
    sections:
      - meeting_link
      - calendar_event
      - meeting_details
      - notes
      - action_items

calendars:
  - id: primary
    name: Primary

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
  assignee: Assignee
  participants: Participants

sync:
  lookahead: today
  statePath: ./.nolendar/state.json
```

### Config Notes

- `microsoft.tenant` supports `common`, `organizations`, or `consumers`
- `microsoft.authMode` supports `device_code`, `interactive_browser`, and `auth_code`
- `calendars` must contain at least one calendar
- `sync.lookahead` defaults to `today`
- `sync.lookahead` also accepts relative ranges like `12h`, `5d`, `2w`, and `3m`
- `sync.statePath` defaults to `.nolendar/state.json`, resolved relative to the config file

## Notion Setup

### Required Meeting Data Source Properties

Nolendar validates these properties on the target meeting data source:

- title property mapped by `mapping.title`
- date property mapped by `mapping.due`
- rich text property mapped by `mapping.eventId`
- rich text property mapped by `mapping.changeKey`
- url property mapped by `mapping.eventLink`, if configured
- multi-select property mapped by `mapping.tags`, if configured and `notion.defaultTags` is non-empty
- people property mapped by `mapping.assignee`, if configured
- relation property mapped by `mapping.participants`, if configured

If `notion.peopleDataSource` is configured, Nolendar also validates that data source for:

- the title property mapped by `notion.peopleDataSource.nameProperty`
- the email property mapped by `notion.peopleDataSource.emailProperty`

If properties are missing, you can create them manually or use `--ensure-properties` where supported.

Notes:

- `--ensure-properties` only applies to auto-creatable properties
- the `Participants` relation must be created in Notion manually and point to the configured People data source

### Property Mapping and Page Content

Nolendar can populate:

- title from the Outlook meeting subject
- due date from the meeting start/end
- event ID from the Outlook event ID
- change key from the Outlook event `changeKey`
- `mapping.eventLink` from Outlook `webLink`
- `mapping.tags` from `notion.defaultTags`
- `mapping.assignee` from `notion.defaultAssigneeEmail` or the authenticated Notion identity
- `mapping.participants` from Outlook attendees through `notion.peopleDataSource`
- `notion.pageIcon` as a static icon on created or updated pages

By default, generated page content includes:

- `Meeting Link`
- `Calendar Event`
- `Meeting Details`
- `Notes`
- `Action items`

You can control generated sections with `notion.pageContent.sections`.

If your template already contains `Notes` or `Action items`, configure only the metadata sections:

```yaml
notion:
  dataSourceTemplate:
    type: default
  pageContent:
    sections:
      - meeting_link
      - calendar_event
      - meeting_details
```

### Template Modes

Nolendar supports two template modes:

- `notion.templatePageId`
  - copies blocks from a normal Notion page and prepends them to Nolendar-generated sections
- `notion.dataSourceTemplate`
  - uses Notion’s native data source template support
  - `sync` creates the page first, then appends Nolendar-generated sections after a short delay
  - `finalize-templates` remains available as a manual fallback

Examples:

```yaml
notion:
  databaseId: your_notion_data_source_id
  templatePageId: 36986680-d5d6-813d-8235-c37663edd559
```

```yaml
notion:
  databaseId: your_notion_data_source_id
  dataSourceTemplate:
    type: default
```

```yaml
notion:
  databaseId: your_notion_data_source_id
  dataSourceTemplate:
    type: template_id
    templateId: a5da15f6-b853-455d-8827-f906fb52db2b
    timezone: America/New_York
```

## Commands

### Setup

Generate a starter config:

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

Validate config:

```bash
npm run dev -- validate-config --config nolendar.yml
```

### Inspection

List meetings:

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

Print the detected Notion schema:

```bash
npm run dev -- print-notion-schema --config nolendar.yml
```

### Sync

Preview sync actions without writing:

```bash
npm run dev -- sync --config nolendar.yml --dry-run
```

Run sync:

```bash
npm run dev -- sync --config nolendar.yml
```

Auto-create missing required properties during sync:

```bash
npm run dev -- sync --config nolendar.yml --ensure-properties
```

Force-update already-synced pages even when the stored Outlook `changeKey` matches:

```bash
npm run dev -- sync --config nolendar.yml --force-update
```

Tune native-template finalize delay:

```bash
npm run dev -- sync --config nolendar.yml --finalize-delay-ms 5000
```

Manual template finalization fallback:

```bash
npm run dev -- finalize-templates --config nolendar.yml
```

### Diagnostics

Show Microsoft Graph timings while listing:

```bash
npm run dev -- list --config nolendar.yml --lookahead 5d --timings
```

Show Notion timings during validation:

```bash
npm run dev -- validate-notion --config nolendar.yml --timings
```

Show both Microsoft Graph and Notion timings during sync:

```bash
npm run dev -- sync --config nolendar.yml --force-update --timings
```

Build and run the compiled CLI:

```bash
npm run build
node dist/index.js list --config nolendar.yml --lookahead 7d
```

## Sync Behavior

The current sync implementation:

- validates the Notion schema before writing
- optionally creates missing required properties with `--ensure-properties`
- filters meetings using:
  - `ignoreDeclined`
  - `minDurationMinutes`
  - `requireAttendees`
  - `ignorePersonal`
  - `ignoreOptionalAttendance`
- looks up existing Notion pages by the configured Outlook event ID property
- skips updates when the stored `changeKey` matches unless `--force-update` is used
- updates existing pages when the `changeKey` differs
- archives matching Notion pages for cancelled meetings that were previously synced
- skips creating new pages for cancelled meetings
- archives matching Notion pages when Graph delta reports deleted events
- collapses multiple delta changes for the same recurring occurrence to the latest final state before sync
- persists delta state only after a successful non-dry-run sync

### Incremental Sync Notes

- `sync` persists per-calendar delta links under `sync.statePath`
- stored delta links are only reused when the saved calendar window exactly matches the current resolved window
- in practice, `today` benefits the most because repeated runs on the same UTC day resolve to the same window
- if the window changes, Nolendar falls back to a fresh `calendarView/delta` bootstrap for that calendar

### Timings

- `--timings` prints one line per outbound Microsoft Graph or Notion API call
- each line includes the service, operation, status, and elapsed duration in milliseconds
- the final `sync` summary includes end-to-end elapsed time

### Current Limitations

Not implemented yet:

- MCP-based Notion integration
- additional auth/token-cache hardening
- helper commands like calendar discovery

## Example Output

Example `list` output:

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

## APIs and Libraries

This project is built around:

- Microsoft Graph API
- Notion API

Libraries:

- [`msgraph-sdk-typescript`](https://github.com/microsoftgraph/msgraph-sdk-typescript)
- [`notion-sdk-js`](https://github.com/makenotion/notion-sdk-js)
- [`@azure/identity`](https://www.npmjs.com/package/@azure/identity)
- [`@azure/msal-node`](https://www.npmjs.com/package/@azure/msal-node)
- [`commander`](https://www.npmjs.com/package/commander)
- [`yaml`](https://www.npmjs.com/package/yaml)

## Development

Development follows TDD: `red -> green -> refactor`.

Common scripts:

```bash
npm test
npm run test:watch
npm run lint
npm run build
```

The implementation plan is tracked in [IMPLEMENTATION_PLAN.md](/Users/rex/workspace/nolendar/IMPLEMENTATION_PLAN.md).
