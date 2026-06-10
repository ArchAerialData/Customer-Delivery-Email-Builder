import uuid
import unittest
from unittest.mock import patch

from sharepoint_excel_sync import (
    COMPLETED_SHEET,
    MASTER_SHEET,
    OIL_GAS_SHEET,
    PROJECT_COLUMNS,
    PROJECT_TABLES,
    ProjectConflictError,
    SharePointProjectDataSource,
    SharePointWorkbookConfig,
    WorkbookFormatError,
)


def values_for(record_id, client, site, modified="2026-01-01T00:00:00Z", modified_by="Remote User"):
    by_column = {
        "Record ID": record_id,
        "Client": client,
        "Project Name": site,
        "Dropbox URLs": f"https://dropbox/{record_id}",
        "Operations Contact Email": f"{record_id}@example.com",
        "Procore": "Procore",
        "Internal URLs": f"https://internal/{record_id}",
        "Last Modified UTC": modified,
        "Last Modified By": modified_by,
    }
    return [by_column[column] for column in PROJECT_COLUMNS]


class FakeGraphWorkbookClient:
    web_url = "https://sharepoint.example/workbook"

    def __init__(self):
        self.columns = {table: list(PROJECT_COLUMNS) for table in PROJECT_TABLES.values()}
        self.rows = {
            PROJECT_TABLES[MASTER_SHEET]: [
                {"index": 0, "values": [values_for("b-record", "Beta Client", "Beta Site", "2026-01-01T00:00:00Z")]},
                {"index": 1, "values": [values_for("a-record", "Alpha Client", "Alpha Site", "2026-01-02T00:00:00Z")]},
            ],
            PROJECT_TABLES[OIL_GAS_SHEET]: [
                {"index": 0, "values": [values_for("oil-record", "Oil Client", "Oil Site", "2026-01-03T00:00:00Z")]},
            ],
            PROJECT_TABLES[COMPLETED_SHEET]: [],
        }
        self.add_calls = []
        self.update_calls = []
        self.delete_calls = []
        self.cache_cleared = False

    def list_tables(self):
        return [{"name": name} for name in self.rows]

    def list_table_columns(self, table_name):
        return [{"name": name} for name in self.columns[table_name]]

    def list_table_rows(self, table_name):
        return [dict(row) for row in self.rows[table_name]]

    def add_table_row(self, table_name, values):
        self.add_calls.append((table_name, values))
        row = {"index": len(self.rows[table_name]), "values": [values]}
        self.rows[table_name].append(row)
        return row

    def update_table_row(self, table_name, index, values):
        self.update_calls.append((table_name, index, values))
        self.rows[table_name][index] = {"index": index, "values": [values]}
        return self.rows[table_name][index]

    def delete_table_row(self, table_name, index):
        self.delete_calls.append((table_name, index))
        del self.rows[table_name][index]
        for next_index, row in enumerate(self.rows[table_name]):
            row["index"] = next_index

    def get_current_user(self):
        return "Unit Tester"

    def clear_token_cache(self):
        self.cache_cleared = True


def make_data_source(client=None):
    return SharePointProjectDataSource(
        SharePointWorkbookConfig(enabled=True, tenant_id="tenant", client_id="client", sharing_url="https://sharepoint.example"),
        client=client or FakeGraphWorkbookClient(),
    )


class SharePointProjectDataSourceTests(unittest.TestCase):
    def test_validate_workbook_passes_when_required_tables_and_columns_exist(self):
        data_source = make_data_source()

        data_source.validate_workbook()

        self.assertEqual(set(data_source._column_cache), set(PROJECT_TABLES.values()))

    def test_validate_workbook_fails_when_required_table_is_missing(self):
        client = FakeGraphWorkbookClient()
        missing_table = PROJECT_TABLES[COMPLETED_SHEET]
        del client.rows[missing_table]
        data_source = make_data_source(client)

        with self.assertRaises(WorkbookFormatError) as ctx:
            data_source.validate_workbook()

        self.assertIn(missing_table, str(ctx.exception))

    def test_validate_workbook_fails_when_required_column_is_missing(self):
        client = FakeGraphWorkbookClient()
        table_name = PROJECT_TABLES[MASTER_SHEET]
        client.columns[table_name] = [column for column in PROJECT_COLUMNS if column != "Record ID"]
        data_source = make_data_source(client)

        with self.assertRaises(WorkbookFormatError) as ctx:
            data_source.validate_workbook()

        self.assertIn(table_name, str(ctx.exception))
        self.assertIn("Record ID", str(ctx.exception))

    def test_list_projects_maps_columns_to_app_fields_and_sorts_by_client_then_site(self):
        data_source = make_data_source()

        rows = data_source.list_projects(MASTER_SHEET)

        self.assertEqual([row["record_id"] for row in rows], ["a-record", "b-record"])
        alpha = rows[0]
        self.assertEqual(alpha["row"], "a-record")
        self.assertEqual(alpha["client"], "Alpha Client")
        self.assertEqual(alpha["site"], "Alpha Site")
        self.assertEqual(alpha["link"], "https://dropbox/a-record")
        self.assertEqual(alpha["email"], "a-record@example.com")
        self.assertEqual(alpha["modified_utc"], "2026-01-02T00:00:00Z")
        self.assertEqual(alpha["source_index"], 1)

    def test_save_new_project_adds_record_id_metadata_and_user(self):
        client = FakeGraphWorkbookClient()
        data_source = make_data_source(client)
        new_uuid = uuid.UUID("12345678-1234-5678-1234-567812345678")

        with patch("sharepoint_excel_sync.uuid.uuid4", return_value=new_uuid):
            data_source.save_project(MASTER_SHEET, {
                "client": "New Client",
                "site": "New Site",
                "link": "https://dropbox/new",
                "email": "new@example.com",
                "procore": "Yes",
                "internal_url": "https://internal/new",
            })

        table_name, values = client.add_calls[-1]
        self.assertEqual(table_name, PROJECT_TABLES[MASTER_SHEET])
        by_column = dict(zip(PROJECT_COLUMNS, values))
        self.assertEqual(by_column["Record ID"], new_uuid.hex)
        self.assertEqual(by_column["Client"], "New Client")
        self.assertEqual(by_column["Project Name"], "New Site")
        self.assertEqual(by_column["Last Modified By"], "Unit Tester")
        self.assertTrue(by_column["Last Modified UTC"].endswith("Z"))

    def test_update_project_uses_record_id_source_index_and_checks_conflict(self):
        client = FakeGraphWorkbookClient()
        data_source = make_data_source(client)

        with self.assertRaises(ProjectConflictError):
            data_source.save_project(
                MASTER_SHEET,
                {"client": "Alpha Client", "site": "Stale Edit"},
                "a-record",
                expected_modified_utc="2025-12-31T00:00:00Z",
            )

        data_source.save_project(
            MASTER_SHEET,
            {
                "client": "Alpha Client",
                "site": "Fresh Edit",
                "link": "https://dropbox/fresh",
                "email": "fresh@example.com",
                "procore": "No",
                "internal_url": "https://internal/fresh",
            },
            "a-record",
            expected_modified_utc="2026-01-02T00:00:00Z",
        )

        table_name, index, values = client.update_calls[-1]
        self.assertEqual(table_name, PROJECT_TABLES[MASTER_SHEET])
        self.assertEqual(index, 1)
        by_column = dict(zip(PROJECT_COLUMNS, values))
        self.assertEqual(by_column["Record ID"], "a-record")
        self.assertEqual(by_column["Project Name"], "Fresh Edit")

    def test_delete_project_checks_conflict_and_force_overrides_it(self):
        client = FakeGraphWorkbookClient()
        data_source = make_data_source(client)

        with self.assertRaises(ProjectConflictError):
            data_source.delete_project(MASTER_SHEET, "b-record", expected_modified_utc="stale")

        data_source.delete_project(MASTER_SHEET, "b-record", expected_modified_utc="stale", force=True)

        self.assertEqual(client.delete_calls[-1], (PROJECT_TABLES[MASTER_SHEET], 0))
        remaining_ids = [row["values"][0][0] for row in client.rows[PROJECT_TABLES[MASTER_SHEET]]]
        self.assertNotIn("b-record", remaining_ids)

    def test_move_project_to_completed_adds_to_completed_then_deletes_source(self):
        client = FakeGraphWorkbookClient()
        data_source = make_data_source(client)

        data_source.move_project_to_completed(
            OIL_GAS_SHEET,
            "oil-record",
            expected_modified_utc="2026-01-03T00:00:00Z",
        )

        self.assertEqual(client.add_calls[-1][0], PROJECT_TABLES[COMPLETED_SHEET])
        completed_values = client.add_calls[-1][1]
        completed_by_column = dict(zip(PROJECT_COLUMNS, completed_values))
        self.assertEqual(completed_by_column["Record ID"], "oil-record")
        self.assertEqual(completed_by_column["Client"], "Oil Client")
        self.assertEqual(client.delete_calls[-1], (PROJECT_TABLES[OIL_GAS_SHEET], 0))

    def test_clear_token_cache_delegates_to_graph_client(self):
        client = FakeGraphWorkbookClient()
        data_source = make_data_source(client)

        data_source.clear_token_cache()

        self.assertTrue(client.cache_cleared)


if __name__ == "__main__":
    unittest.main()
