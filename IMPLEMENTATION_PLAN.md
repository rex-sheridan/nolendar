# Nolendar Implementation Plan

## Goal

Build a command-line tool that reads upcoming Outlook calendar meetings via Microsoft Graph and creates or updates corresponding Notion pages without creating duplicates.

## Delivery Principles

- Use TDD throughout: write failing tests first, then implement, then refactor.
- Optimize for idempotent sync behavior before adding convenience features.
- Keep the first release CLI-first and config-driven.
- Separate provider integrations from sync logic so Notion API and MCP support can evolve independently.

## Proposed Tech Stack

- Language: TypeScript
- Runtime: Node.js
- Package manager: npm or pnpm
- Test runner: Vitest
- CLI framework: `commander` or `yargs`
- Config parser: `yaml`
- Microsoft Graph client: `msgraph-sdk-typescript`
- Notion client: `notion-sdk-js`

## High-Level Architecture

Core modules:

- `config`
  - Load, validate, and normalize YAML config.
- `cli`
  - Parse command-line flags and dispatch commands.
- `graph`
  - Authenticate and fetch calendar events and delta tokens from Microsoft Graph.
- `notion`
  - Query database, validate schema, create pages, update pages, apply template behavior.
- `sync`
  - Compare Outlook events to existing Notion pages and decide create, update, skip, or archive behavior.
- `mapping`
  - Convert normalized meeting data into Notion property payloads.
- `filters`
  - Apply config-driven event filtering rules.
- `storage`
  - Persist sync state such as delta links/tokens per calendar.
- `domain`
  - Shared types for meetings, sync records, config, and Notion mapping.

## Data Model

### Normalized Meeting

Represent Graph events in an internal format that includes:

- Outlook event ID
- `changeKey`
- Calendar ID
- Subject
- Start/end datetime with timezone
- Organizer
- Required attendees
- Optional attendees
- Meeting URL
- Event URL/backlink
- Body preview or agenda content
- Response status
- Recurrence metadata
- Is personal/private flag if available

### Required Notion Database Properties

Minimum required properties:

- `Name` or configured title field
- `Due`
- `Outlook Event ID`
- `Outlook ChangeKey`

Recommended properties:

- Organizer
- Attendees
- Meeting Link
- Status
- Notes
- Action Items
- Calendar ID or calendar name

### Local Sync State

Persist state in a local file, for example under `.nolendar/state.json`, keyed by calendar:

- Calendar identifier
- Delta token or delta link
- Last successful sync time
- Optional schema validation cache

## CLI Scope

Initial commands:

- `nolendar sync`
  - Fetch meetings in the configured lookahead window and create/update Notion pages.
- `nolendar list`
  - Print upcoming meetings for the requested time period.
- `nolendar validate-config`
  - Validate YAML config and required secrets.
- `nolendar validate-notion`
  - Check database access and required properties.
- `nolendar init`
  - Generate a starter YAML config.

Useful flags:

- `--config <path>`
- `--calendar <id>`
- `--lookahead today|24h|7d`
- `--dry-run`
- `--verbose`

## Configuration Design

YAML config should include:

- Microsoft account and auth settings
- Target Notion database ID
- Optional Notion template/page template configuration
- One or more calendars to sync
- Default lookahead window
- Event filters
- Field mappings from meeting fields to Notion properties
- State storage path

Example sections:

- `microsoft`
- `notion`
- `calendars`
- `filters`
- `mapping`
- `sync`

## Authentication Plan

### Microsoft Graph

Support OAuth flows that work for both personal and work Microsoft accounts.

Phase 1 recommendation:

- Device code flow or local interactive auth for developer use
- Token caching on disk

Later improvement:

- Non-interactive auth options where feasible

### Notion

- Use Notion integration token for API mode
- Add MCP adapter later if API parity is insufficient or MCP use is preferred

## Sync Algorithm

### First Sync

1. Load config and validate prerequisites.
2. Validate Notion database schema.
3. Fetch meetings for each configured calendar in the lookahead window.
4. Apply filtering rules.
5. Normalize meetings into the internal model.
6. Query Notion for existing pages by `Outlook Event ID`.
7. For each event:
   - Create a page if no page exists.
   - Update the page if `changeKey` differs.
   - Skip if the event already exists and `changeKey` matches.
8. Save delta token/link for each calendar after a successful sync.

### Incremental Sync

1. Load saved delta token/link for a calendar.
2. Fetch changed events using Graph delta queries.
3. Re-run filter and mapping logic on returned changes.
4. Create, update, or mark deleted/cancelled meetings appropriately.
5. Persist the new delta token/link only after the full calendar sync succeeds.

### Deletions and Cancellations

Decide early how to handle removed meetings:

- Preferred initial behavior: update the Notion page status to cancelled rather than delete the page.
- If Graph reports deletions, locate the page by `Outlook Event ID` and update status or archive it.

## Recurring Meetings

Requirements:

- Treat each occurrence as a separate syncable meeting instance.
- Use the actual occurrence event ID returned by Graph, not only the series master.
- Preserve idempotency for each occurrence.
- Verify delta behavior for recurring edits, moved instances, and cancellations.

This area needs dedicated tests before implementation because recurrence bugs will create duplicate or stale pages quickly.

## Notion Template Strategy

Phase 1:

- Create pages with a standardized property set and optional body content sections for notes and action items.

Phase 2:

- If a template page is configured, duplicate or emulate its structure when creating new meeting pages.

Need to confirm the exact Notion API constraints around template application before locking implementation details.

## TDD Roadmap

### Phase 0: Project Setup

Write tests for:

- Config loading and validation
- CLI argument parsing
- Environment/secrets validation

Implement:

- TypeScript project scaffold
- Vitest setup
- Lint/format config
- Basic CLI entrypoint

### Phase 1: Config and Domain Model

Write tests for:

- YAML parsing
- Defaults for lookahead behavior
- Invalid config rejection
- Field mapping normalization
- Filter config evaluation

Implement:

- Config schema
- Domain types
- Filter engine

### Phase 2: Meeting Listing

Write tests for:

- Graph event normalization
- Time window selection
- Multiple calendar aggregation
- CLI output formatting

Implement:

- Graph client wrapper
- Meeting list command

### Phase 3: Notion Schema Validation

Write tests for:

- Detection of missing required properties
- Property creation flow or user prompt behavior
- Mapping validation against database schema

Implement:

- Notion database inspector
- Required property validator

### Phase 4: Initial Sync

Write tests for:

- Create page for unseen meeting
- Skip page when matching `Outlook Event ID` and `changeKey`
- Update page when `changeKey` changes
- Apply filter rules before page creation
- Set `Due` property correctly

Implement:

- Sync engine
- Notion page creation/update
- Idempotency checks

### Phase 5: Incremental Sync

Write tests for:

- Delta token persistence
- Changed event updates
- Cancelled/deleted event handling
- Partial failure behavior does not corrupt saved sync state

Implement:

- Delta query support
- State store
- Transaction-like sync completion semantics

### Phase 6: Recurrence and Edge Cases

Write tests for:

- Recurring series with multiple occurrences
- Single occurrence moved or edited
- Single occurrence cancelled
- Meetings with no attendees
- Declined meetings
- Short-duration meetings

Implement:

- Recurrence-safe sync handling
- Edge-case filtering refinements

### Phase 7: Template and UX Improvements

Write tests for:

- Template-enabled page creation
- Dry-run output
- Verbose sync summaries

Implement:

- Template support
- Better CLI messaging
- Operational diagnostics

## Testing Strategy

Test layers:

- Unit tests for config, filtering, mapping, and sync decisions
- Contract tests for Graph and Notion adapters using mocked SDK clients
- Integration tests around sync flows with fixture events and fixture Notion records

Critical test focus:

- Idempotency
- Delta token correctness
- Recurring meeting handling
- Mapping correctness
- Cancellation/deletion behavior

## Error Handling

Expected failure categories:

- Invalid config
- Missing credentials
- Graph auth failure
- Notion permission or schema mismatch
- Network/API rate limiting
- Partial sync failure

Approach:

- Fail fast on config and schema problems
- Emit actionable CLI errors
- Do not advance delta state if sync is incomplete
- Support `--dry-run` for safe verification

## Release Plan

### Milestone 1

- Project scaffold
- Config loader
- `list` command

### Milestone 2

- Notion schema validation
- Initial sync with idempotent create/update

### Milestone 3

- Incremental sync with delta queries
- Local sync state persistence

### Milestone 4

- Recurring meeting correctness
- Filters and template support

## Open Decisions

- Whether Notion MCP support should ship in the first release or after direct API support
- Whether missing Notion properties should be auto-created by default or only after explicit confirmation
- Whether cancelled meetings should be archived or marked with a status field
- Exact OAuth flow and token cache approach for personal and enterprise Microsoft accounts
- The minimum supported Node.js version

## Recommended First Build Order

1. Scaffold the TypeScript CLI project with Vitest.
2. Implement config parsing and validation.
3. Implement the `list` command against Microsoft Graph.
4. Implement Notion database validation and required-property checks.
5. Implement initial sync with strict idempotency.
6. Add delta-query incremental sync.
7. Harden recurrence handling and edge cases.
