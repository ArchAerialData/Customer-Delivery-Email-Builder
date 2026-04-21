import base64
import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

import msal
import requests

logger = logging.getLogger(__name__)
RESERVED_SCOPES = {"openid", "profile", "offline_access"}
DEFAULT_SCOPES = ["User.Read", "Mail.ReadWrite"]


class GraphEmailClient:
    """MSAL-backed Microsoft Graph helper that creates Draft messages (no send)."""

    def __init__(self, settings_path: Path, cache_path: Path) -> None:
        self._settings_path = Path(settings_path)
        self._token_cache_path = Path(cache_path)
        self._settings = self._load_settings()
        self._cache = msal.SerializableTokenCache()
        if self._token_cache_path.exists():
            try:
                self._cache.deserialize(self._token_cache_path.read_text(encoding="utf-8"))
            except Exception:
                self._cache = msal.SerializableTokenCache()
                self._token_cache_path.unlink(missing_ok=True)
        self._app = msal.PublicClientApplication(
            client_id=self._settings["client_id"],
            authority=self._settings["authority"],
            token_cache=self._cache,
        )

    def _load_settings(self) -> Dict[str, Any]:
        if not self._settings_path.exists():
            raise FileNotFoundError(f"Missing Graph settings: {self._settings_path}")
        data = json.loads(self._settings_path.read_text(encoding="utf-8"))
        required = ["client_id", "tenant_id", "authority"]
        for key in required:
            if not data.get(key):
                raise ValueError(f"graph_app_settings.json missing required key: {key}")
        scopes = data.get("scopes") or DEFAULT_SCOPES
        if not isinstance(scopes, list):
            scopes = DEFAULT_SCOPES
        filtered = [s for s in scopes if s not in RESERVED_SCOPES]
        if len(filtered) != len(scopes):
            logger.warning("Removed reserved scopes from graph_app_settings.json: %s", RESERVED_SCOPES)
        if not filtered:
            filtered = DEFAULT_SCOPES
        data["scopes"] = filtered
        data["redirect_uri"] = data.get("redirect_uri") or None
        return data

    def _save_cache(self) -> None:
        if self._cache.has_state_changed:
            self._token_cache_path.write_text(self._cache.serialize(), encoding="utf-8")

    def reset_cache(self) -> None:
        self._cache = msal.SerializableTokenCache()
        self._token_cache_path.unlink(missing_ok=True)
        self._app.token_cache = self._cache

    def acquire_token(self) -> Dict[str, Any]:
        scopes = self._settings.get("scopes", DEFAULT_SCOPES)
        accounts = self._app.get_accounts()
        result: Optional[Dict[str, Any]] = None
        if accounts:
            result = self._app.acquire_token_silent(scopes, account=accounts[0])
        if not result:
            kwargs: Dict[str, Any] = {"scopes": scopes, "prompt": "select_account"}
            result = self._app.acquire_token_interactive(**kwargs)
        if "access_token" not in result:
            raise RuntimeError(f"Failed to acquire Graph token: {result.get('error_description')}")
        self._save_cache()
        return result

    @staticmethod
    def _recipient(addr: str) -> Dict[str, Any]:
        return {"emailAddress": {"address": addr}}

    @staticmethod
    def _file_attachment(path: Path, inline: bool = False, cid: Optional[str] = None) -> Dict[str, Any]:
        content = path.read_bytes()
        att: Dict[str, Any] = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": path.name,
            "contentBytes": base64.b64encode(content).decode("ascii"),
        }
        if inline:
            att["isInline"] = True
            att["contentId"] = cid or path.stem
        return att

    def create_draft(
        self,
        subject: str,
        body_html: str,
        to: List[str],
        cc: List[str],
        bcc: List[str],
        attachments: List[Path],
        inline_attachments: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        token = self.acquire_token()
        message: Dict[str, Any] = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [self._recipient(a) for a in to],
            "ccRecipients": [self._recipient(a) for a in cc],
            "bccRecipients": [self._recipient(a) for a in bcc],
            "attachments": [],
        }

        for path in attachments:
            message["attachments"].append(self._file_attachment(path))

        if inline_attachments:
            for att in inline_attachments:
                if "contentBytes" in att:
                    message["attachments"].append(att)
                    continue
                p = Path(att["path"])
                cid = att.get("cid") or p.stem
                message["attachments"].append(self._file_attachment(p, inline=True, cid=cid))

        resp = requests.post(
            "https://graph.microsoft.com/v1.0/me/messages",
            headers={"Authorization": f"Bearer {token['access_token']}"},
            json=message,
            timeout=30,
        )
        if not resp.ok:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text
            raise RuntimeError(f"Graph create draft failed ({resp.status_code}): {detail}")
        return resp.json()
