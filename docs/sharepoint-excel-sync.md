# SharePoint Excel Sync Implementation Guide

## Current Implementation

The app now has a SharePoint-ready project data layer while preserving the local workbook fallback.

- Local mode still reads and writes `Reference Data/XLSX Workbooks/Customer Information.xlsx`.
- SharePoint mode is enabled only when a per-user `sharepoint_config.json` has `"enabled": true`.
- SharePoint mode uses Microsoft Graph delegated auth with MSAL, resolves the configured sharing URL, creates persistent workbook sessions, validates required Excel Tables, and performs row-level add/update/delete operations.
- The project tabs use the same Add/Edit/Delete/Reload/Move To Completed controls in both modes.
- SharePoint conflict checks compare `Last Modified UTC` before updating, deleting, or moving a row.
- The old local byte-level undo is disabled in SharePoint mode; use SharePoint version history for rollback.

This is a v1 implementation of the recommended architecture. It does not add webhook push sync, app-only background auth, or whole-file overwrite sync.

## Required Workbook Shape

Before using SharePoint mode, prepare and upload a workbook with these Excel Tables:

- `ConstructionProjects` on `Construction`
- `OilGasProjects` on `Oil & Gas`
- `CompletedProjects` on `Completed Projects`

Each table must include these columns:

- `Record ID`
- `Client`
- `Project Name`
- `Dropbox URLs`
- `Operations Contact Email`
- `Procore`
- `Internal URLs`
- `Last Modified UTC`
- `Last Modified By`

The local workbook can be converted into a SharePoint-ready copy:

```bash
python scripts/prepare_sharepoint_workbook.py
```

The default output is:

```text
Reference Data/XLSX Workbooks/Customer Information.sharepoint-ready.xlsx
```

Upload that prepared copy to SharePoint and test against the copy first. Do not point production users at the live workbook until the probe and pilot pass.

## App Registration

Register a Microsoft Entra desktop/public-client app.

Required app-registration settings:

- Single-tenant app registration for the organization tenant.
- Platform: `Mobile and desktop applications`.
- Redirect URI for system-browser desktop auth: `http://localhost`.
- Public client flows: enabled.

Recommended delegated scopes:

- `User.Read`
- `Files.ReadWrite`
- `offline_access`

Do not use a client secret in the desktop app. Do not use username/password auth.

Tenant policy may still block user consent even though the delegated `Files.ReadWrite` permission does not inherently require admin consent. Test with a normal team member account before rollout.

## Configuration

The app loads config from the first matching location:

1. The file pointed to by `EMAIL_BUILDER_SHAREPOINT_CONFIG`
2. Windows: `%APPDATA%\Email Builder\sharepoint_config.json`
3. macOS: `~/Library/Application Support/Email Builder/sharepoint_config.json`
4. Linux/other: `~/.config/email-builder/sharepoint_config.json`

Start from `sharepoint_config.example.json`, but place the real config in a per-user config folder rather than committing it.

Example:

```json
{
  "enabled": true,
  "tenant_id": "your-tenant-id",
  "client_id": "your-public-client-app-id",
  "sharing_url": "https://your-tenant.sharepoint.com/:x:/s/site-name/...",
  "scopes": [
    "User.Read",
    "Files.ReadWrite",
    "offline_access"
  ]
}
```

MSAL token cache is stored beside the config as `.sharepoint_token_cache.json`.

## Probe Workflow

Install dependencies:

```bash
python -m pip install openpyxl ttkbootstrap msal requests
```

Validate auth, sharing URL resolution, workbook sessions, required tables, and row reads:

```bash
python scripts/sharepoint_workbook_probe.py
```

Against a sandbox workbook only, validate row add/update/move/delete:

```bash
python scripts/sharepoint_workbook_probe.py --write-test
```

The write test creates a temporary row, updates it, moves it to `Completed Projects`, then deletes it.

## Runtime Behavior

Local mode:

- Uses the existing local workbook.
- Sorts the local workbook after edits.
- Keeps one local undo snapshot.

SharePoint mode:

- Uses Graph table row APIs.
- Reads table rows with `$top`/`$skip` paging so larger workbooks do not depend on one oversized response.
- Retries retryable Graph responses such as throttling/timeouts once, honoring `Retry-After` when present.
- Recreates the persistent workbook session once if Graph reports that the previous session expired.
- Displays sorted rows in the app without changing the workbook's current sort/filter state.
- Writes immediately on Save/Delete/Move.
- Prompts before overwriting/deleting/moving a row that changed remotely since the edit dialog or shortcut-tab selection loaded.
- Adds an `Open SharePoint` button to project tabs.
- Adds a `Sign Out` button in SharePoint mode to clear the local MSAL token cache.
- Disables local undo because rollback belongs to SharePoint version history.

## Validation Checklist

Before team rollout:

- A normal user can sign in without admin intervention.
- The sharing URL resolves through Graph.
- The probe sees the expected workbook name and web URL.
- All three tables are discovered by name.
- All required columns are present.
- Manual Excel Online edits appear in the app after Reload.
- App edits appear in Excel Online after Save.
- Two-user conflicts produce a prompt instead of silent overwrite.
- The app opens cleanly when SharePoint is unavailable and reports a useful status.
- The prepared workbook has been piloted as a SharePoint copy, not directly on production data.

## Known Caveats

- Excel Graph table row APIs require delegated user auth; app-only permissions are not supported for this v1 row workflow.
- Webhooks are not included because a local desktop app would need a public HTTPS notification endpoint.
- The CC-list workbook remains local in v1.
- Whole-file replacement is intentionally avoided to reduce overwrite risk.
- If users rename/delete tables or required columns in Excel Online, SharePoint mode will fail validation until the workbook is repaired.
