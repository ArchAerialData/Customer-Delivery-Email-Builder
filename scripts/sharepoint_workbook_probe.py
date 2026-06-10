#!/usr/bin/env python3
"""Probe the configured SharePoint workbook through Microsoft Graph.

Use this before enabling SharePoint mode for the desktop app:

    python scripts/sharepoint_workbook_probe.py
    python scripts/sharepoint_workbook_probe.py --write-test

The script uses the same per-user config file as the app unless
EMAIL_BUILDER_SHAREPOINT_CONFIG points to a different JSON file.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from sharepoint_excel_sync import (
    MASTER_SHEET,
    PROJECT_TABLES,
    SharePointProjectDataSource,
    SharePointWorkbookConfig,
)


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe SharePoint workbook access for Email Builder.")
    parser.add_argument(
        "--write-test",
        action="store_true",
        help="Add, update, move, and delete a temporary test row. Use only against a sandbox workbook.",
    )
    args = parser.parse_args()

    config = SharePointWorkbookConfig.load()
    if not config.enabled:
        print(f"SharePoint sync is not enabled in config: {config.config_path}")
        return 2

    data_source = SharePointProjectDataSource(config)
    item = data_source.client.resolve_drive_item()
    print("Resolved workbook")
    print(f"  name: {item.get('name')}")
    print(f"  webUrl: {item.get('webUrl')}")
    print(f"  driveId: {(item.get('parentReference') or {}).get('driveId')}")
    print(f"  itemId: {item.get('id')}")

    data_source.ensure_schema()
    print("Workbook schema is valid.")

    worksheets = data_source.client.list_worksheets()
    tables = data_source.client.list_tables()
    print("Worksheets:", ", ".join(str(ws.get("name")) for ws in worksheets))
    print("Tables:", ", ".join(str(table.get("name")) for table in tables))

    for sheet_name, table_name in PROJECT_TABLES.items():
        rows = data_source.list_projects(sheet_name, allow_missing=True)
        print(f"{sheet_name} / {table_name}: {len(rows)} project row(s)")

    if args.write_test:
        print("Running write test against sandbox workbook...")
        values = {
            "client": "ZZZ Graph Probe",
            "site": "Temporary Test Project",
            "link": "https://example.invalid/dropbox",
            "email": "noreply@example.invalid",
            "procore": "Probe",
            "internal_url": "https://example.invalid/internal",
        }
        data_source.save_project(MASTER_SHEET, values)
        created = next(
            row for row in data_source.list_projects(MASTER_SHEET)
            if row["client"] == values["client"] and row["site"] == values["site"]
        )
        values["site"] = "Temporary Test Project Updated"
        data_source.save_project(
            MASTER_SHEET,
            values,
            created["row"],
            expected_modified_utc=created.get("modified_utc"),
        )
        updated = next(row for row in data_source.list_projects(MASTER_SHEET) if row["row"] == created["row"])
        data_source.move_project_to_completed(
            MASTER_SHEET,
            updated["row"],
            expected_modified_utc=updated.get("modified_utc"),
        )
        moved = next(row for row in data_source.list_projects("Completed Projects") if row["row"] == created["row"])
        data_source.delete_project(
            "Completed Projects",
            moved["row"],
            expected_modified_utc=moved.get("modified_utc"),
        )
        print("Write test completed.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
