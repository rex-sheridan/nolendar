# nolendar

  <p align="center">
    <img src="docs/assets/nolendar-logo.png" alt="Nolendar logo" width="180">
  </p>

Nolendar reads upcoming meetings from Outlook Calendar via Microsoft Graph and turns them into structured Notion pages for notes, action items, and related meeting context.

See [docs/architecture.md](docs/architecture.md) for Mermaid architecture and process-flow diagrams.

## Current Status

Implemented now:

- Outlook meeting listing from one or more Microsoft calendars
- Outlook calendar discovery with copyable calendar IDs
- Notion page creation and update with idempotency based on Outlook event ID and `changeKey`
- Incremental sync using Microsoft Graph `calendarView/delta`
- Deleted-event archiving and recurring-meeting handling
- Notion template support via:
  - `notion.templatePageId`
  - `notion.dataSourceTemplate`
- Participant relation syncing through a separate Notion People data source
- Configurable filters, tags, assignee resolution, page icons, and page content sections
- General Markdown augmentation of matching Notion meeting pages under a chosen heading
- CLI timings for Microsoft Graph and Notion API calls

Remaining work is mostly polish:

- optional MCP-based Notion integration

## Quickstart

1. Install dependencies:

```bash
npm install
```

2. Make the local CLI available as `nolendar`:

```bash
export PATH="$PATH:$(pwd)/bin"
```

To keep it available in new shells, add that line to `~/.zshrc`, or use an alias:

```bash
alias nolendar="$(pwd)/bin/nolendar"
```

3. Run the setup wizard:

```bash
nolendar wizard
```

   The wizard writes progress to `$XDG_CONFIG_HOME/nolendar/config.yml` (default `~/.config/nolendar/config.yml`) as you go. It walks through Microsoft authentication, Outlook calendar selection, Notion data source selection, property mappings, and optional sections such as filters, templates, page content, and sync settings.

   You can also run the same interactive flow through `init`:

```bash
nolendar init --wizard
```

   Before running the wizard, follow [AUTHENTICATION.md](AUTHENTICATION.md) to provide the required `MICROSOFT_CLIENT_ID` and `NOTION_TOKEN` or `NOTION_API_KEY`.

4. If you prefer to edit YAML manually, generate a starter config instead:

```bash
nolendar init
```

5. Validate config:

```bash
nolendar validate-config
```

6. Validate Notion access and schema:

```bash
nolendar validate-notion
```

7. Preview sync without writing:

```bash
nolendar sync --dry-run
```

8. Run sync:

```bash
nolendar sync
```

9. Augment today's Notion meeting pages with Markdown under any heading:

```bash
nolendar augment --input augment.md --heading "AI Preparation" --dry-run
nolendar augment --input augment.md --heading "AI Preparation"
```

The input file should put each meeting title on its own line, followed by the Markdown content to append to that meeting:

```markdown
Planning
- Key decision: confirm launch scope.
- Before the meeting: review open risks.

Design Review
## Decisions to drive
- Approve or reject the final layout direction.
```

Matching is scoped to the selected day and uses the Notion meeting title after case and whitespace normalization. Use `--day YYYY-MM-DD` for a different day. Omit `--input` to read from stdin, for example `agent-command | nolendar augment --heading "Follow-ups"`.

If your generated input needs explicit meeting separators, set `notion.augmentation.delimiter` in the config file or pass `--delimiter <text>`. The CLI flag overrides the config value. The delimiter is matched as a full line after trimming whitespace; the first recognized meeting title after each delimiter becomes the section target:

```bash
nolendar augment --input augment.md --heading "AI Preparation" --delimiter "%%MEETING%%"
```

```yaml
notion:
  augmentation:
    delimiter: "%%MEETING%%"
```

```markdown
%%MEETING%%
Planning
- Key decision: confirm launch scope.

%%MEETING%%
Design Review
- Approve or reject the final layout direction.
```

## Requirements

- Node.js 20+
- access to the Outlook calendars you want to read
- a Notion token with access to the target data source

Detailed auth and credential setup lives in [AUTHENTICATION.md](/Users/rex/workspace/nolendar/AUTHENTICATION.md).

## Configuration

Nolendar follows the [XDG Base Directory Specification](https://specifications.freedesktop.org/basedir-spec/latest/) and does not discover application files from the current working directory:

- config: `$XDG_CONFIG_HOME/nolendar/config.yml` (default `~/.config/nolendar/config.yml`)
- environment: `$XDG_CONFIG_HOME/nolendar/env` (default `~/.config/nolendar/env`)
- Microsoft token cache: `$XDG_DATA_HOME/nolendar/msal-cache.json` (default `~/.local/share/nolendar/msal-cache.json`)
- sync state: `$XDG_STATE_HOME/nolendar/state.json` (default `~/.local/state/nolendar/state.json`)

Create the optional environment file for local credentials with:

```bash
mkdir -p "${XDG_CONFIG_HOME:-$HOME/.config}/nolendar"
cp .env.example "${XDG_CONFIG_HOME:-$HOME/.config}/nolendar/env"
```

Shell-exported variables take precedence over values in the environment file. `--config <path>`, `sync.statePath`, and `MICROSOFT_TOKEN_CACHE_PATH` remain explicit overrides.

### Migrating a legacy installation

From the checkout or installation directory that contains the old `nolendar.yml`, `.env`, and `.nolendar/`, preview the migration:

```bash
bin/migrate-to-xdg --dry-run
```

Then copy the legacy files into their XDG locations:

```bash
bin/migrate-to-xdg
```

Use `--legacy-dir PATH` when invoking the script from another directory. The script removes only the generated legacy `sync.statePath: .nolendar/state.json` setting from the copied config, applies private file permissions, refuses to overwrite conflicting destinations, and leaves every source file in place for rollback. It is safe to run again after a successful migration.

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
  augmentation:
    delimiter: "%%MEETING%%"
  peopleDataSource:
    databaseId: your_people_data_source_id
    nameProperty: Name
    emailProperty: Email Address
    maxAttendeesPerMeeting: 10
  canceledMeetings:
    action: set_status
    statusProperty: Status
    statusValue: Canceled
  completedMeetings:
    statusProperty: Status
    doneStatusValue: Done
    canceledStatusValue: Canceled
    lookback: 1d
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
  ignoreNames:
    - Focus time
  ignorePatterns:
    - "^OOO"

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
```

### Config Notes

- `microsoft.tenant` supports `common`, `organizations`, or `consumers`
- `microsoft.authMode` supports `device_code`, `interactive_browser`, and `auth_code`
- use `npm run dev -- list-calendars` to print available Outlook calendar IDs before choosing `calendars[].id`
- Microsoft Graph auth requires either `MICROSOFT_CLIENT_ID` for delegated public-client auth or `MICROSOFT_ACCESS_TOKEN` for a raw-token override
- `calendars` must contain at least one calendar
- `sync.lookahead` defaults to `today`
- `sync.lookahead` also accepts relative ranges like `12h`, `5d`, `2w`, and `3m`
- `sync.statePath` defaults to the XDG state file; an explicit relative value is resolved relative to the config file
- `filters.ignoreNames` skips events whose title exactly matches one of the configured strings
- `filters.ignorePatterns` skips events whose title matches one of the configured JavaScript regular expressions
- `notion.peopleDataSource.maxAttendeesPerMeeting` defaults to `10`; set it lower to reduce Notion People lookups for large meetings, or `0` to skip participant relation association while still creating meeting pages
- `notion.canceledMeetings.action` defaults to `archive`; set it to `set_status` with `statusProperty` and `statusValue` to mark matching pages, for example `Status: Canceled`
- `notion.completedMeetings` marks synced Notion entries whose meeting date has passed as done; when canceled meetings use `set_status`, it defaults to the same status property, `Done`, the configured canceled status value, and a `1d` lookback before the active sync window

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
- status property mapped by `notion.canceledMeetings.statusProperty`, if `notion.canceledMeetings.action` is `set_status`
- status property mapped by `notion.completedMeetings.statusProperty`, if `notion.completedMeetings` is configured

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

When participant syncing is enabled, Nolendar associates at most `notion.peopleDataSource.maxAttendeesPerMeeting` attendees per meeting with the configured relation. This limits expensive Notion People lookup/create calls for large meetings.

By default, generated page content includes:

- `Meeting Link`
- `Calendar Event`
- `Meeting Details`
- `Notes`
- `Action items`

You can control generated sections with `notion.pageContent.sections`.
To place them within a template instead of after all template content, set
`notion.pageContent.insertAfterHeading` to the exact text of a heading block.
Nolendar keeps that marker heading, inserts its generated blocks immediately
after it, and keeps the remaining template blocks below the generated content.
The `augment` command uses the same marker when inserting imported Markdown.
If the configured heading is missing, the operation fails with a clear error.

If your template already contains `Notes` or `Action items`, configure only the metadata sections:

```yaml
notion:
  dataSourceTemplate:
    type: default
  pageContent:
    insertAfterHeading: Nolendar Content
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

Run the interactive setup wizard:

```bash
npm run dev -- wizard
```

Run the wizard through `init`:

```bash
npm run dev -- init --wizard
```

Write wizard output to a custom path:

```bash
npm run dev -- wizard --config ./config/nolendar.yml
```

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

List available Outlook calendars:

```bash
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
npm run dev -- list-calendars
```

Use Microsoft auth settings from a config file:

```bash
npm run dev -- list-calendars --config nolendar.yml
```

Override Microsoft auth settings without a config file:

```bash
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
npm run dev -- list-calendars --tenant organizations --auth-mode interactive_browser
```

List meetings:

```bash
npm run dev -- list --config nolendar.yml
```

Override the lookahead window:

```bash
npm run dev -- list --config nolendar.yml --lookahead 5d
```

Print AI-readable meeting contents for a single day from Outlook:

```bash
npm run dev -- meetings --config nolendar.yml
```

Inspect another day or read synced Notion pages instead:

```bash
npm run dev -- meetings --config nolendar.yml --day tomorrow --source notion
```

Include all normalized properties with each meeting:

```bash
npm run dev -- meetings --config nolendar.yml --source notion --full-properties
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

Shorten long IDs in timing and verbose sync logs:

```bash
npm run dev -- sync --config nolendar.yml --timings --verbose --compact-ids
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
- archives matching Notion pages for cancelled meetings that were previously synced, unless `notion.canceledMeetings.action` is `set_status`
- sets the configured status value on matching Notion pages for cancelled meetings when `notion.canceledMeetings.action` is `set_status`
- skips creating new pages for cancelled meetings
- archives matching Notion pages when Graph delta reports deleted events, or sets the configured status when `notion.canceledMeetings.action` is `set_status`
- sets the configured done status on synced Notion meeting pages whose meeting date has passed, unless they are already done or canceled
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
- add `--compact-ids` to shorten long IDs as `first...last` in timing output and verbose sync decisions
- the final `sync` summary includes end-to-end elapsed time

### Current Limitations

Not implemented yet:

- MCP-based Notion integration
- additional auth/token-cache hardening

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
