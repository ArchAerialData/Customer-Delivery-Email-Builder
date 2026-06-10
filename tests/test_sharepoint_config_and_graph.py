import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import sharepoint_excel_sync as sp


class FakeResponse:
    def __init__(self, status_code, payload=None, text="", headers=None):
        self.status_code = status_code
        self._payload = payload
        self.text = text
        self.headers = headers or {}
        self.content = b"" if payload is None else json.dumps(payload).encode("utf-8")

    def json(self):
        if self._payload is None:
            raise ValueError("No JSON")
        return self._payload


class FakeRequests:
    def __init__(self, responses):
        self.responses = list(responses)
        self.calls = []

    def request(self, method, url, **kwargs):
        self.calls.append((method, url, kwargs))
        return self.responses.pop(0)


class FakeTokenProvider:
    def acquire_token(self):
        return {"access_token": "unit-token"}


class SharePointConfigAndGraphTests(unittest.TestCase):
    def test_encode_sharing_url_round_trips_to_original_url(self):
        url = "https://example.sharepoint.com/:x:/s/team/File.xlsx?e=abc123"
        token = sp.encode_sharing_url(url)

        self.assertTrue(token.startswith("u!"))
        padding = "=" * (-len(token[2:]) % 4)
        decoded = base64.urlsafe_b64decode(token[2:] + padding).decode("utf-8")
        self.assertEqual(decoded, url)

    def test_config_load_defaults_to_disabled_when_file_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_path = Path(tmp) / "missing.json"
            config = sp.SharePointWorkbookConfig.load(config_path)

        self.assertFalse(config.enabled)
        self.assertEqual(config.config_path, config_path)
        self.assertEqual(config.token_cache_path, config_path.with_name(".sharepoint_token_cache.json"))

    def test_config_load_trims_values_and_does_not_treat_string_false_as_enabled(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_path = Path(tmp) / "sharepoint_config.json"
            token_path = Path(tmp) / "token.json"
            config_path.write_text(json.dumps({
                "enabled": "false",
                "tenant_id": " tenant ",
                "client_id": " client ",
                "sharing_url": " https://sharepoint.example/workbook ",
                "scopes": [" User.Read ", "", "Files.ReadWrite"],
                "token_cache_path": str(token_path),
            }), encoding="utf-8")

            config = sp.SharePointWorkbookConfig.load(config_path)

        self.assertFalse(config.enabled)
        self.assertEqual(config.tenant_id, "tenant")
        self.assertEqual(config.client_id, "client")
        self.assertEqual(config.sharing_url, "https://sharepoint.example/workbook")
        self.assertEqual(config.scopes, ["User.Read", "Files.ReadWrite"])
        self.assertEqual(config.token_cache_path, token_path)

    def test_enabled_config_validation_requires_core_fields(self):
        config = sp.SharePointWorkbookConfig(enabled=True, tenant_id="", client_id="abc", sharing_url="")

        with self.assertRaises(sp.SharePointSyncError) as ctx:
            config.validate()

        self.assertIn("tenant_id", str(ctx.exception))
        self.assertIn("sharing_url", str(ctx.exception))

    def test_graph_error_extracts_status_code_error_code_and_retry_after(self):
        response = FakeResponse(
            429,
            payload={"error": {"code": "TooManyRequests", "message": "Slow down"}},
            headers={"Retry-After": "7"},
        )

        error = sp._graph_error(response)

        self.assertEqual(error.status_code, 429)
        self.assertEqual(error.code, "TooManyRequests")
        self.assertEqual(error.retry_after, 7)
        self.assertIn("Slow down", str(error))

    def test_graph_request_retries_retryable_status_and_keeps_session_header(self):
        client = object.__new__(sp.GraphWorkbookClient)
        client._requests = FakeRequests([
            FakeResponse(429, payload={"error": {"code": "TooManyRequests", "message": "retry"}}, headers={"Retry-After": "0"}),
            FakeResponse(200, payload={"ok": True}),
        ])
        client._token_provider = FakeTokenProvider()
        client._session_id = "session-123"

        with patch.object(sp.time, "sleep") as sleep:
            result = client._request("GET", "/workbook/test", workbook_session=True)

        self.assertEqual(result, {"ok": True})
        self.assertEqual(len(client._requests.calls), 2)
        self.assertEqual(client._requests.calls[0][2]["headers"]["Workbook-Session-Id"], "session-123")
        sleep.assert_called_once_with(0)

    def test_graph_request_refreshes_expired_workbook_session_once(self):
        client = object.__new__(sp.GraphWorkbookClient)
        client._requests = FakeRequests([
            FakeResponse(404, payload={"error": {"code": "invalidSession", "message": "expired"}}),
            FakeResponse(201, payload={"id": "fresh-session"}),
            FakeResponse(200, payload={"ok": True}),
        ])
        client._token_provider = FakeTokenProvider()
        client._drive_item = {"id": "item-id", "parentReference": {"driveId": "drive-id"}}
        client._session_id = "expired-session"

        result = client._request("GET", "/workbook/tables", workbook_session=True)

        self.assertEqual(result, {"ok": True})
        self.assertEqual(client._session_id, "fresh-session")
        methods_and_urls = [(method, url) for method, url, _kwargs in client._requests.calls]
        self.assertEqual(methods_and_urls, [
            ("GET", "https://graph.microsoft.com/v1.0/workbook/tables"),
            ("POST", "https://graph.microsoft.com/v1.0/drives/drive-id/items/item-id/workbook/createSession"),
            ("GET", "https://graph.microsoft.com/v1.0/workbook/tables"),
        ])
        self.assertEqual(client._requests.calls[0][2]["headers"]["Workbook-Session-Id"], "expired-session")
        self.assertEqual(client._requests.calls[2][2]["headers"]["Workbook-Session-Id"], "fresh-session")

    def test_graph_paged_values_follows_next_links_until_exhausted(self):
        client = object.__new__(sp.GraphWorkbookClient)
        seen = []

        def fake_request(method, url, **kwargs):
            seen.append((method, url, kwargs))
            if url == "/first":
                return {"value": [{"name": "one"}], "@odata.nextLink": "https://graph.example/next"}
            return {"value": [{"name": "two"}]}

        client._request = fake_request

        values = sp.GraphWorkbookClient._paged_values(client, "/first")

        self.assertEqual(values, [{"name": "one"}, {"name": "two"}])
        self.assertEqual([call[1] for call in seen], ["/first", "https://graph.example/next"])
        self.assertTrue(all(call[2]["workbook_session"] for call in seen))

    def test_list_table_rows_uses_top_skip_paging_when_next_links_are_absent(self):
        client = object.__new__(sp.GraphWorkbookClient)
        client.ensure_session = lambda: None
        client._table_path = lambda table_name, suffix="": f"/tables/{table_name}{suffix}"
        seen_params = []

        def fake_request(method, url, **kwargs):
            seen_params.append(kwargs.get("params"))
            skip = (kwargs.get("params") or {}).get("$skip", 0)
            if skip == 0:
                return {"value": [{"index": idx} for idx in range(500)]}
            return {"value": [{"index": 500}]}

        client._request = fake_request

        rows = sp.GraphWorkbookClient.list_table_rows(client, "Table1")

        self.assertEqual(len(rows), 501)
        self.assertEqual(seen_params, [{"$top": 500}, {"$top": 500, "$skip": 500}])

    def test_resolve_drive_item_uses_encoded_share_url_and_caches_response(self):
        config = sp.SharePointWorkbookConfig(
            enabled=True,
            tenant_id="tenant",
            client_id="client",
            sharing_url="https://sharepoint.example/:x:/s/team/workbook.xlsx?e=abc",
        )
        client = object.__new__(sp.GraphWorkbookClient)
        client._config = config
        client._drive_item = None
        calls = []

        def fake_request(method, path, **kwargs):
            calls.append((method, path, kwargs))
            return {"id": "item-id", "parentReference": {"driveId": "drive-id"}}

        client._request = fake_request

        first = sp.GraphWorkbookClient.resolve_drive_item(client)
        second = sp.GraphWorkbookClient.resolve_drive_item(client)

        self.assertIs(first, second)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], "GET")
        self.assertTrue(calls[0][1].startswith("/shares/u!"))
        self.assertTrue(calls[0][1].endswith("/driveItem"))

    def test_create_session_uses_persistent_session_endpoint_and_stores_id(self):
        client = object.__new__(sp.GraphWorkbookClient)
        client._drive_item = {"id": "item-id", "parentReference": {"driveId": "drive-id"}}
        client._session_id = None
        calls = []

        def fake_request(method, path, **kwargs):
            calls.append((method, path, kwargs))
            return {"id": "session-id"}

        client._request = fake_request

        session_id = sp.GraphWorkbookClient.create_session(client, persist_changes=True)

        self.assertEqual(session_id, "session-id")
        self.assertEqual(client._session_id, "session-id")
        self.assertEqual(calls[0][0], "POST")
        self.assertEqual(calls[0][1], "/drives/drive-id/items/item-id/workbook/createSession")
        self.assertEqual(calls[0][2]["json_body"], {"persistChanges": True})
        self.assertEqual(calls[0][2]["expected"], (201,))

    def test_clear_token_cache_removes_disk_cache_and_replaces_in_memory_cache(self):
        class FakeMsal:
            @staticmethod
            def SerializableTokenCache():
                return {"fresh": True}

        class FakeApp:
            token_cache = {"stale": True}

        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / ".sharepoint_token_cache.json"
            cache_path.write_text("old-cache", encoding="utf-8")
            provider = object.__new__(sp.GraphTokenProvider)
            provider._msal = FakeMsal
            provider._app = FakeApp()
            provider._cache_path = cache_path

            provider.clear_cache()

            self.assertFalse(cache_path.exists())
            self.assertEqual(provider._cache, {"fresh": True})
            self.assertIs(provider._app.token_cache, provider._cache)


if __name__ == "__main__":
    unittest.main()
