"""
One-time interactive sign-in to mint a refresh token for the dispatch app.

Run this ONCE, signed in as the sending mailbox (isabelle@practus.com). It
opens a browser, Isabelle consents to Mail.Send + Mail.ReadWrite, and the
resulting refresh token is printed. Paste that into backend/.env as
GRAPH_REFRESH_TOKEN and the server can then acquire access tokens silently.

    python backend/get_refresh_token.py

Prerequisite — on the app registration, under Authentication, add a
"Web" platform with this exact redirect URI:

    http://localhost:8400/callback

and under API permissions add DELEGATED Microsoft Graph permissions
Mail.Send and Mail.ReadWrite.
"""

import os
import sys
import socket
import logging
import secrets
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import msal
from dotenv import load_dotenv

BACKEND_DIR = Path(__file__).resolve().parent
load_dotenv(BACKEND_DIR / ".env")

REDIRECT_PORT = int(os.environ.get("GRAPH_AUTH_PORT", "8400"))
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}/callback"
SCOPES = ["Mail.Send", "Mail.ReadWrite"]

logging.basicConfig(level=logging.WARNING)

# Filled in by the callback handler.
_result: dict = {}


class _CallbackHandler(BaseHTTPRequestHandler):
    """Catches the single redirect back from Microsoft."""

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/callback":
            self.send_response(404)
            self.end_headers()
            return

        params = urllib.parse.parse_qs(parsed.query)
        _result.update({k: v[0] for k, v in params.items()})

        ok = "code" in _result
        title = "Sign-in complete" if ok else "Sign-in failed"
        detail = (
            "You can close this tab and return to the terminal."
            if ok
            else f"{_result.get('error', 'unknown')}: {_result.get('error_description', '')}"
        )
        body = (
            "<html><body style=\"font-family:system-ui;padding:40px\">"
            f"<h2>{title}</h2><p>{detail}</p></body></html>"
        ).encode("utf-8")

        self.send_response(200 if ok else 400)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass  # keep the console clean


def _port_is_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def main() -> int:
    tenant_id = os.environ.get("GRAPH_TENANT_ID", "").strip()
    client_id = os.environ.get("GRAPH_CLIENT_ID", "").strip()
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "").strip()

    missing = [
        n
        for n, v in (
            ("GRAPH_TENANT_ID", tenant_id),
            ("GRAPH_CLIENT_ID", client_id),
            ("GRAPH_CLIENT_SECRET", client_secret),
        )
        if not v
    ]
    if missing:
        print(f"ERROR: missing in backend/.env: {', '.join(missing)}")
        return 1

    if not _port_is_free(REDIRECT_PORT):
        print(f"ERROR: port {REDIRECT_PORT} is in use. Free it, or set GRAPH_AUTH_PORT")
        print("       (and add the matching redirect URI in Azure).")
        return 1

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )

    state = secrets.token_urlsafe(16)
    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        state=state,
        prompt="select_account",
    )

    print("=" * 72)
    print("Sign in as the SENDING mailbox (isabelle@practus.com).")
    print("A browser window should open. If it does not, paste this URL:")
    print()
    print(auth_url)
    print("=" * 72)

    server = HTTPServer(("127.0.0.1", REDIRECT_PORT), _CallbackHandler)
    webbrowser.open(auth_url)
    print(f"Waiting for the redirect on {REDIRECT_URI} ...")
    server.handle_request()  # blocks until Microsoft redirects back
    server.server_close()

    if "code" not in _result:
        print(f"\nAuthorization failed: {_result.get('error')}")
        print(_result.get("error_description", ""))
        return 1

    if _result.get("state") != state:
        print("\nERROR: state mismatch - aborting (possible interference).")
        return 1

    token = app.acquire_token_by_authorization_code(
        _result["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )

    if "refresh_token" not in token:
        print(f"\nToken exchange failed: {token.get('error')}")
        print((token.get("error_description") or "")[:500])
        if "access_token" in token:
            print("\nAn access token came back but no refresh token. Ensure the app")
            print("registration allows offline_access for this account.")
        return 1

    account = (token.get("id_token_claims") or {}).get("preferred_username", "unknown")
    granted = token.get("scope", "")

    print()
    print("=" * 72)
    print(f"Signed in as : {account}")
    print(f"Scopes       : {granted}")
    print("=" * 72)
    print()
    print("Add this line to backend/.env (single line, no quotes):")
    print()
    print(f"GRAPH_REFRESH_TOKEN={token['refresh_token']}")
    print()
    print("Treat it like a password: it grants mailbox access until revoked.")
    print("=" * 72)
    return 0


if __name__ == "__main__":
    sys.exit(main())
