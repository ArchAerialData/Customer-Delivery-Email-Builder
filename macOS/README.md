# macOS Support

These files are for running the app on a Mac as a Customer Information workbook editor. On macOS, email draft creation is disabled.

## GitHub Build Workflow

The macOS package is built by `.github/workflows/build-macos-app.yml`.

The macOS and Windows package workflows are temporarily disabled while other features are being developed. This prevents package builds from running on every commit.

While disabled:

- `.github/workflows/build-macos-app.yml` is manual-only and its job is skipped with `if: ${{ false }}`
- `.github/workflows/build-exe.yml` is manual-only and its job is skipped with `if: ${{ false }}`

To re-enable package builds:

1. Remove the `if: ${{ false }}` guard from the workflow job you want to run.
2. Restore the desired automatic triggers.
3. For macOS branch builds, restore `push` and `pull_request` triggers for the `MacOS-Compatibility` branch in `.github/workflows/build-macos-app.yml`.
4. For Windows release builds, restore the `push` tag trigger for `v*` tags in `.github/workflows/build-exe.yml`.

The workflow installs the macOS Python dependencies, validates the source resources, builds `Email Builder.app` with PyInstaller, smoke-tests the packaged app, creates `dist/Email Builder.dmg`, verifies the DMG with `hdiutil`, and uploads it as the `email-builder-macos` artifact.

## Install From DMG

Download the `email-builder-macos` artifact from the GitHub Actions run, open `Email Builder.dmg`, and drag `Email Builder.app` to `Applications`.

This app is not signed or notarized. The first launch will likely be blocked by Gatekeeper.

To open it:

1. Try to open `Email Builder.app`.
2. When macOS blocks it, open `System Settings`.
3. Select `Privacy & Security` in the sidebar.
4. Scroll to the bottom of the `Privacy & Security` tab.
5. Click `Open Anyway` for `Email Builder.app`.
6. Confirm that you want to open the app.

Packaged macOS builds copy their starter data into:

```bash
~/Library/Application Support/Email Builder/Reference Data
```

Existing files in that folder are not overwritten, so user-edited workbook data is preserved.

## Run From Source

Install Python 3 from python.org, then install the app dependencies:

```bash
python3 -m pip install -r macOS/requirements-macos.txt
```

## Run

Open `Run Email Builder.command`, or run this from the project folder:

```bash
python3 email_template_gui.py
```

If macOS says the command file is not executable, run:

```bash
chmod +x "macOS/Run Email Builder.command"
```

When running from source, the workbook editing screens use `Reference Data/XLSX Workbooks/Customer Information.xlsx`.

## Optional SharePoint Workbook Sync

The app can use a SharePoint-hosted `Customer Information.xlsx` workbook when a per-user SharePoint config is enabled. Local workbook mode remains the default.

Before enabling SharePoint mode:

1. Register a Microsoft Entra public-client desktop app with delegated `User.Read`, `Files.ReadWrite`, and `offline_access` scopes.
2. Prepare a SharePoint-ready workbook copy:

```bash
python3 scripts/prepare_sharepoint_workbook.py
```

3. Upload the prepared copy to SharePoint.
4. Create `~/Library/Application Support/Email Builder/sharepoint_config.json` from `sharepoint_config.example.json` and set `"enabled": true`.
5. Validate access against the SharePoint copy:

```bash
python3 scripts/sharepoint_workbook_probe.py
```

Use `scripts/sharepoint_workbook_probe.py --write-test` only against a sandbox workbook. More details are in `docs/sharepoint-excel-sync.md`.
