import unittest
from unittest.mock import patch

import email_template_gui as gui
from sharepoint_excel_sync import ProjectConflictError


class FakeStatusVar:
    def __init__(self):
        self.value = ""

    def set(self, value):
        self.value = value


class FakeRemoteDataSource:
    is_remote = True
    web_url = "https://sharepoint.example/workbook"

    def __init__(self, conflict_once=False):
        self.conflict_once = conflict_once
        self.save_calls = []
        self.delete_calls = []
        self.move_calls = []
        self.cleared = False

    def save_project(self, *args, **kwargs):
        self.save_calls.append((args, kwargs))
        if self.conflict_once and not kwargs.get("force"):
            self.conflict_once = False
            raise ProjectConflictError("changed remotely")

    def delete_project(self, *args, **kwargs):
        self.delete_calls.append((args, kwargs))
        if self.conflict_once and not kwargs.get("force"):
            self.conflict_once = False
            raise ProjectConflictError("changed remotely")

    def move_project_to_completed(self, *args, **kwargs):
        self.move_calls.append((args, kwargs))
        if self.conflict_once and not kwargs.get("force"):
            self.conflict_once = False
            raise ProjectConflictError("changed remotely")

    def clear_token_cache(self):
        self.cleared = True

    def list_projects(self, *args, **kwargs):
        return [{
            "row": "record-123",
            "client": "Client",
            "site": "Site",
            "link": "https://dropbox.example",
            "email": "client@example.com",
            "modified_utc": "2026-01-01T00:00:00Z",
        }]


class FakeLocalDataSource:
    is_remote = False

    def __init__(self):
        self.save_calls = []

    def save_project(self, *args, **kwargs):
        self.save_calls.append((args, kwargs))


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class FakeTree:
    def __init__(self, selection):
        self._selection = selection

    def selection(self):
        return self._selection


def make_app(data_source, error=""):
    app = object.__new__(gui.EmailTemplateApp)
    app.project_data_source = data_source
    app._project_data_source_error = error
    app._undo_stack = []
    app.status_var = FakeStatusVar()
    app.reloads = []
    app.status_messages = []
    app.undo_paths = []
    app._reload_master_tab = lambda sheet_name: app.reloads.append(sheet_name)
    app._reload_master_data = lambda *args, **kwargs: None
    app._set_sync_status = lambda message=None: app.status_messages.append(message)
    app._record_undo = lambda path: app.undo_paths.append(path)
    return app


class EmailTemplateSharePointIntegrationTests(unittest.TestCase):
    def test_sync_status_reports_remote_unavailable_when_sharepoint_error_exists(self):
        app = make_app(FakeRemoteDataSource(), error="missing table")

        status = gui.EmailTemplateApp._sync_status_text(app)

        self.assertIn("SharePoint workbook unavailable", status)
        self.assertIn("missing table", status)

    def test_read_master_data_clears_stale_sharepoint_error_after_successful_remote_read(self):
        app = make_app(FakeRemoteDataSource(), error="previous failure")

        rows = gui.EmailTemplateApp._read_master_data(app, gui.MASTER_SHEET)

        self.assertEqual(rows[0]["row"], "record-123")
        self.assertEqual(app._project_data_source_error, "")

    def test_save_master_entry_passes_row_identity_and_expected_modified_to_remote_source(self):
        data_source = FakeRemoteDataSource()
        app = make_app(data_source)
        row = {"row": "record-123", "modified_utc": "2026-01-01T00:00:00Z"}
        values = {"client": "Client", "site": "Site"}

        gui.EmailTemplateApp._save_master_entry(app, gui.MASTER_SHEET, values, row)

        args, kwargs = data_source.save_calls[-1]
        self.assertEqual(args, (gui.MASTER_SHEET, values, "record-123"))
        self.assertEqual(kwargs["expected_modified_utc"], "2026-01-01T00:00:00Z")
        self.assertEqual(app.reloads, [gui.MASTER_SHEET])
        self.assertEqual(app.status_messages, ["Saved project data."])

    def test_save_master_entry_conflict_confirm_retries_with_force(self):
        data_source = FakeRemoteDataSource(conflict_once=True)
        app = make_app(data_source)
        row = {"row": "record-123", "modified_utc": "old"}
        values = {"client": "Client", "site": "Site"}

        with patch.object(gui.messagebox, "askyesno", return_value=True):
            gui.EmailTemplateApp._save_master_entry(app, gui.MASTER_SHEET, values, row)

        self.assertEqual(len(data_source.save_calls), 2)
        self.assertFalse(data_source.save_calls[0][1].get("force", False))
        self.assertTrue(data_source.save_calls[1][1]["force"])
        self.assertEqual(app.reloads, [gui.MASTER_SHEET])

    def test_save_master_entry_conflict_decline_does_not_reload(self):
        data_source = FakeRemoteDataSource(conflict_once=True)
        app = make_app(data_source)

        with patch.object(gui.messagebox, "askyesno", return_value=False):
            with self.assertRaises(ProjectConflictError):
                gui.EmailTemplateApp._save_master_entry(
                    app,
                    gui.MASTER_SHEET,
                    {"client": "Client"},
                    {"row": "record-123", "modified_utc": "old"},
                )

        self.assertEqual(len(data_source.save_calls), 1)
        self.assertEqual(app.reloads, [])

    def test_delete_and_move_conflict_confirm_use_force(self):
        delete_source = FakeRemoteDataSource(conflict_once=True)
        delete_app = make_app(delete_source)
        move_source = FakeRemoteDataSource(conflict_once=True)
        move_app = make_app(move_source)
        row = {"row": "record-123", "modified_utc": "old"}

        with patch.object(gui.messagebox, "askyesno", return_value=True):
            gui.EmailTemplateApp._delete_project_row_with_conflict_prompt(delete_app, gui.MASTER_SHEET, row)
            gui.EmailTemplateApp._move_project_row_with_conflict_prompt(move_app, gui.MASTER_SHEET, row)

        self.assertTrue(delete_source.delete_calls[-1][1]["force"])
        self.assertTrue(move_source.move_calls[-1][1]["force"])

    def test_local_save_records_undo_before_writing(self):
        data_source = FakeLocalDataSource()
        app = make_app(data_source)

        gui.EmailTemplateApp._save_master_entry(app, gui.MASTER_SHEET, {"client": "Local"}, None)

        self.assertEqual(app.undo_paths, [gui.MASTER_FILE])
        self.assertEqual(len(data_source.save_calls), 1)

    def test_sharepoint_undo_shows_version_history_message_without_local_restore(self):
        app = make_app(FakeRemoteDataSource())

        with patch.object(gui.messagebox, "showinfo") as showinfo:
            gui.EmailTemplateApp._undo_last_change(app)

        self.assertIn("version history", app.status_var.value)
        showinfo.assert_called_once()

    def test_clear_sharepoint_token_cache_delegates_to_data_source_after_confirmation(self):
        data_source = FakeRemoteDataSource()
        app = make_app(data_source)

        with patch.object(gui.messagebox, "askyesno", return_value=True):
            gui.EmailTemplateApp._clear_sharepoint_token_cache(app)

        self.assertTrue(data_source.cleared)
        self.assertEqual(app.status_messages, ["SharePoint sign-in token cleared. Sign in again on the next sync."])

    def test_dropbox_update_uses_sharepoint_conflict_prompt_and_force_retry(self):
        data_source = FakeRemoteDataSource(conflict_once=True)
        app = make_app(data_source)
        app.sheet_data = {
            gui.MASTER_SHEET: [{
                "row": "record-123",
                "client": "Old Client",
                "site": "Old Site",
                "link": "old-link",
                "email": "old@example.com",
                "procore": "old-procore",
                "internal_url": "old-internal",
                "modified_utc": "old",
            }]
        }
        app.dropbox_tree = FakeTree(["record-123"])
        app.db_client_var = FakeVar("New Client")
        app.db_site_var = FakeVar("New Site")
        app.db_link_var = FakeVar("new-link")
        app.dropbox_reloads = 0
        app._reload_dropbox_data = lambda: setattr(app, "dropbox_reloads", app.dropbox_reloads + 1)

        with patch.object(gui.messagebox, "askyesno", return_value=True):
            gui.EmailTemplateApp._update_dropbox_entry(app)

        self.assertEqual(len(data_source.save_calls), 2)
        self.assertTrue(data_source.save_calls[-1][1]["force"])
        self.assertEqual(app.dropbox_reloads, 1)

    def test_contact_update_uses_sharepoint_conflict_prompt_and_force_retry(self):
        data_source = FakeRemoteDataSource(conflict_once=True)
        app = make_app(data_source)
        app.sheet_data = {
            gui.MASTER_SHEET: [{
                "row": "record-123",
                "client": "Old Client",
                "site": "Old Site",
                "link": "old-link",
                "email": "old@example.com",
                "procore": "old-procore",
                "internal_url": "old-internal",
                "modified_utc": "old",
            }]
        }
        app.contacts_tree = FakeTree(["record-123"])
        app.ct_client_var = FakeVar("New Client")
        app.ct_site_var = FakeVar("New Site")
        app.ct_email_var = FakeVar("new@example.com")
        app.contact_reloads = 0
        app._reload_contacts_data = lambda: setattr(app, "contact_reloads", app.contact_reloads + 1)

        with patch.object(gui.messagebox, "askyesno", return_value=True):
            gui.EmailTemplateApp._update_contact_entry(app)

        self.assertEqual(len(data_source.save_calls), 2)
        self.assertTrue(data_source.save_calls[-1][1]["force"])
        self.assertEqual(app.contact_reloads, 1)


if __name__ == "__main__":
    unittest.main()
