# Authentication and Credentials

This document covers:

- Microsoft Graph authentication modes
- environment variables
- Notion token setup
- Outlook calendar IDs
- Notion data source IDs
- the usual live-run checklist

## Environment Variables

Microsoft Graph auth may use:

```bash
export MICROSOFT_CLIENT_ID=your_app_registration_client_id
```

For `device_code` and `interactive_browser`, `MICROSOFT_CLIENT_ID` is required. Use an app registration configured for public-client delegated auth.

For `microsoft.authMode: auth_code`, you also need:

```bash
export MICROSOFT_CLIENT_SECRET=your_client_secret
export MICROSOFT_REDIRECT_URI=http://localhost:8787/auth/callback
```

You can also bypass Nolendar’s built-in Microsoft auth flows entirely by supplying a raw bearer token:

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

## Microsoft Authentication Modes

### Option 1: `device_code`

Use this if your app registration can enable public client flows.

High-level steps:

1. Open `https://entra.microsoft.com/`
2. Go to `Microsoft Entra ID` -> `Identity` -> `Applications` -> `App registrations`
3. Select `New registration`
4. Enter an app name such as `nolendar`
5. Choose the supported account type
6. Leave `Redirect URI` empty
7. Create the app registration
8. Open `Authentication`
9. Under `Advanced settings`, set `Allow public client flows` to `Yes`
10. Copy the `Application (client) ID` and export it as `MICROSOFT_CLIENT_ID`

Config:

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

Notes:

- device-code flow does not require a client secret
- Nolendar uses delegated Graph scopes for user sign-in
- Microsoft will prompt you to open `https://microsoft.com/devicelogin`

### Option 2: `auth_code`

Use this if your app registration is configured as a web app and you can create a client secret.

High-level steps:

1. Open the app registration in Microsoft Entra
2. Go to `Authentication`
3. Add or confirm a `Web` platform
4. Add a redirect URI such as `http://localhost:8787/auth/callback`
5. Save
6. Go to `Certificates & secrets`
7. Create a new client secret
8. Copy the `Application (client) ID`
9. Copy the secret value immediately after creating it

Config:

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

- `MICROSOFT_REDIRECT_URI` must exactly match the app registration
- Nolendar opens a browser and waits for the localhost callback

### Option 3: `interactive_browser`

Use this when your app registration is configured as a public client and you want browser-based sign-in.

Config:

```yaml
microsoft:
  tenant: common
  authMode: interactive_browser
```

Environment:

```bash
export MICROSOFT_GRAPH_SCOPES=Calendars.Read,User.Read
export MICROSOFT_CLIENT_ID=your_microsoft_app_client_id
```

Notes:

- the app registration must be usable as a public client
- Nolendar opens a browser for sign-in

### Option 4: `MICROSOFT_ACCESS_TOKEN` override

Use this when you already have a short-lived Graph bearer token from another client and want Nolendar to use it directly.

Environment:

```bash
export MICROSOFT_ACCESS_TOKEN=eyJ...
```

Notes:

- this takes precedence over Nolendar’s configured auth mode
- it is useful as a developer fallback
- the token is short-lived and must be refreshed manually
- Graph Explorer is a practical source for this token

## Notion Token Setup

1. Create a Notion integration in the Notion developer dashboard
2. Copy the integration token
3. Share the target Notion data source with that integration
4. Export the token:

```bash
export NOTION_TOKEN=secret_xxx
```

## Outlook Calendar IDs

For a first live run, you can usually start with:

```yaml
calendars:
  - id: primary
    name: Primary
```

If you want a non-primary calendar, list the calendars available to the signed-in Microsoft account:

```bash
npm run dev -- list-calendars
```

If you already have Microsoft auth settings in `nolendar.yml`, use them:

```bash
npm run dev -- list-calendars --config nolendar.yml
```

Example output:

```text
Available calendars:
  - Calendar [default]
    id: primary-calendar-id
  - Team
    id: team-calendar-id
    owner: Owner <owner@example.com>
```

Copy the `id` value into `calendars[].id` in `nolendar.yml`.

For large meetings, you can cap how many attendees are associated through the Notion People relation:

```yaml
notion:
  peopleDataSource:
    databaseId: your_people_data_source_id
    nameProperty: Name
    emailProperty: Email Address
    maxAttendeesPerMeeting: 10
```

Set `maxAttendeesPerMeeting` to `0` to create meeting pages without participant relation association.

## Notion Data Source ID

`notion.databaseId` should be the target Notion data source ID used for querying and page creation.

If you only have the parent database/container ID, you may still need the underlying data source ID from Notion’s data source management UI.

## Live Run Checklist

Once your values are set, the usual sequence is:

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
