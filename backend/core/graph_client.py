"""
Microsoft Graph API client — sends emails as Isabelle using DELEGATED auth.

Uses the OAuth refresh-token flow: a one-time interactive sign-in (see
get_refresh_token.py) produces a long-lived refresh token, which this client
exchanges for short-lived access tokens silently, with no user interaction.

Delegated permissions required on the app registration: Mail.Send, Mail.ReadWrite.

Environment variables:
  GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, GRAPH_REFRESH_TOKEN,
  GRAPH_SENDER_EMAIL, GRAPH_SENDER_NAME, GRAPH_DRY_RUN
"""

import os
import time
import logging
from functools import lru_cache
from pathlib import Path
from uuid import uuid4

import msal
import requests
from dotenv import load_dotenv

# Load backend/.env regardless of the process working directory.
load_dotenv(Path(__file__).resolve().parent.parent / ".env")

log = logging.getLogger("graph_client")

DRY_RUN_PREFIX = "dry-run-"

# Delegated scopes. MSAL adds openid/profile/offline_access itself — passing
# those reserved values here raises, so list only the resource scopes.
SCOPES = ["Mail.Send", "Mail.ReadWrite"]

# Refresh an access token this many seconds before it actually expires.
EXPIRY_SKEW = 300


def _truthy(val: str | None) -> bool:
    return str(val or "").strip().lower() in {"1", "true", "yes", "on"}


class GraphAuthError(RuntimeError):
    """Raised when a refresh token cannot be exchanged for an access token."""


class GraphClient:
    def __init__(self):
        self.tenant_id = os.environ.get("GRAPH_TENANT_ID", "")
        self.client_id = os.environ.get("GRAPH_CLIENT_ID", "")
        self.client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "")
        self.sender_email = os.environ.get("GRAPH_SENDER_EMAIL", "")
        self.sender_name = os.environ.get("GRAPH_SENDER_NAME") or self.sender_email

        self._refresh_token = os.environ.get("GRAPH_REFRESH_TOKEN", "").strip()
        self._token: str | None = None
        self._expires_at: float = 0.0
        self._app: msal.ConfidentialClientApplication | None = None

        # Without full credentials (or with GRAPH_DRY_RUN=1) nothing is sent to
        # Microsoft: calls are logged and fake ids returned, so the parse and
        # preview flow can be exercised end to end.
        self.dry_run = _truthy(os.environ.get("GRAPH_DRY_RUN")) or not self.configured
        if self.dry_run:
            missing = ", ".join(self.missing_vars) or "none"
            log.warning(f"GraphClient in DRY-RUN mode - no mail will leave Outlook (missing: {missing})")

    # -- Configuration -------------------------------------

    @property
    def missing_vars(self) -> list[str]:
        return [
            name
            for name, val in (
                ("GRAPH_TENANT_ID", self.tenant_id),
                ("GRAPH_CLIENT_ID", self.client_id),
                ("GRAPH_CLIENT_SECRET", self.client_secret),
                ("GRAPH_REFRESH_TOKEN", self._refresh_token),
                ("GRAPH_SENDER_EMAIL", self.sender_email),
            )
            if not val.strip()
        ]

    @property
    def configured(self) -> bool:
        return not self.missing_vars

    # -- Delegated auth (refresh token -> access token) -----

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}"

    def _msal_app(self) -> msal.ConfidentialClientApplication:
        if self._app is None:
            self._app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=self.authority,
            )
        return self._app

    def _get_token(self) -> str:
        """Exchange the refresh token for a fresh access token."""
        result = self._msal_app().acquire_token_by_refresh_token(
            self._refresh_token, scopes=SCOPES
        )

        if "access_token" not in result:
            err = result.get("error", "unknown_error")
            desc = (result.get("error_description") or "")[:300]
            if err in {"invalid_grant", "interaction_required"}:
                raise GraphAuthError(
                    f"Refresh token rejected ({err}). Re-run get_refresh_token.py "
                    f"and update GRAPH_REFRESH_TOKEN. Details: {desc}"
                )
            raise GraphAuthError(f"Token request failed ({err}): {desc}")

        # Entra rotates the refresh token on each use; keep the newest one for
        # the life of this process so long-running servers do not go stale.
        rotated = result.get("refresh_token")
        if rotated and rotated != self._refresh_token:
            self._refresh_token = rotated
            log.info("Refresh token rotated (in-memory only - .env still holds the original)")

        self._token = result["access_token"]
        self._expires_at = time.time() + int(result.get("expires_in", 3600))
        log.debug(f"Access token acquired, valid for {result.get('expires_in')}s")
        return self._token

    @property
    def token(self) -> str:
        if not self._token or time.time() >= self._expires_at - EXPIRY_SKEW:
            return self._get_token()
        return self._token

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
        }

    def _call(self, method: str, path: str, json_body: dict | None = None) -> requests.Response:
        url = f"https://graph.microsoft.com/v1.0{path}"
        r = requests.request(method, url, headers=self._headers(), json=json_body)
        if r.status_code == 401:
            self._token = None  # force a fresh exchange, not the cached token
            r = requests.request(method, url, headers=self._headers(), json=json_body)
        return r

    def whoami(self) -> dict:
        """Signed-in user behind the refresh token - handy for verifying setup."""
        r = self._call("GET", "/me")
        return r.json() if r.status_code == 200 else {"error": r.status_code, "detail": r.text[:200]}

    # -- Mail ----------------------------------------------
    # Delegated auth acts as the signed-in user, so these use /me rather than
    # /users/{address}: that needs no directory-read permission.

    def create_draft(
        self, to: list[str], cc: list[str], subject: str, body_html: str
    ) -> str | None:
        """Create a draft email. Returns message_id or None."""
        payload = {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body_html},
            "toRecipients": [{"emailAddress": {"address": e}} for e in to],
            "ccRecipients": [{"emailAddress": {"address": e}} for e in cc],
            "isDraft": True,
        }
        if self.dry_run:
            fake_id = DRY_RUN_PREFIX + uuid4().hex[:12]
            log.info(f"[dry-run] draft -> to={to} cc={cc} subject={subject!r} id={fake_id}")
            return fake_id

        r = self._call("POST", "/me/messages", payload)
        if r.status_code == 201:
            return r.json()["id"]
        log.error(f"Draft failed: {r.status_code} {r.text[:200]}")
        return None

    def send_draft(self, message_id: str) -> bool:
        """Send an existing draft."""
        if self.dry_run or message_id.startswith(DRY_RUN_PREFIX):
            log.info(f"[dry-run] send draft {message_id} - not actually sent")
            return True
        r = self._call("POST", f"/me/messages/{message_id}/send")
        if r.status_code != 202:
            log.error(f"Send failed: {r.status_code} {r.text[:200]}")
        return r.status_code == 202

    def delete_draft(self, message_id: str) -> bool:
        """Delete a draft."""
        if self.dry_run or message_id.startswith(DRY_RUN_PREFIX):
            log.info(f"[dry-run] delete draft {message_id}")
            return True
        r = self._call("DELETE", f"/me/messages/{message_id}")
        return r.status_code == 204

    def send_email(
        self, to: list[str], cc: list[str], subject: str, body_html: str
    ) -> bool:
        """Send immediately (no draft)."""
        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": e}} for e in to],
                "ccRecipients": [{"emailAddress": {"address": e}} for e in cc],
            }
        }
        if self.dry_run:
            log.info(f"[dry-run] sendMail -> to={to} cc={cc} subject={subject!r}")
            return True

        r = self._call("POST", "/me/sendMail", payload)
        if r.status_code != 202:
            log.error(f"sendMail failed: {r.status_code} {r.text[:200]}")
        return r.status_code == 202


@lru_cache()
def get_graph_client() -> GraphClient:
    return GraphClient()
