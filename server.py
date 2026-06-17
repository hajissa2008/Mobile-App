#!/usr/bin/env python3
"""Proxy server for SharePoint Activity Dashboard."""

import http.server
import json
import urllib.parse
import webbrowser
import requests
import msal

SHAREPOINT_HOST = "sascontractingsa.sharepoint.com"
SHAREPOINT_SITE = "/sites/Alenshaiah"
SHAREPOINT_LIST = "Maseel Activity"

CLIENT_ID = "1950a258-227b-4e31-a9cf-717495945fc2"
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPE = [f"https://{SHAREPOINT_HOST}/.default"]

PORT = 5000

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
token_cache = None


def get_token():
    global token_cache
    try:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPE, account=accounts[0])
            if result and "access_token" in result:
                token_cache = result
                return result["access_token"]
    except Exception as e:
        print(f"[TOKEN] Silent acquisition failed: {e}")
    if token_cache and "access_token" in token_cache:
        return token_cache["access_token"]
    return None


def start_device_login():
    flow = app.initiate_device_flow(SCOPE)
    if "user_code" not in flow:
        return {"error": "Failed to create device flow", "details": flow.get("error_description", "")}
    return {
        "user_code": flow["user_code"],
        "device_code": flow["device_code"],
        "verification_uri": flow["verification_uri"],
        "message": flow["message"],
        "expires_in": flow["expires_in"],
    }


def poll_token(device_code):
    global token_cache
    result = app.acquire_token_by_device_flow({"device_code": device_code})
    if "access_token" in result:
        token_cache = result
        accounts = app.get_accounts()
        print(f"[AUTH] Login successful. Access token acquired. Accounts: {len(accounts)}")
        return {"status": "success", "access_token": result["access_token"][:20] + "..."}
    error = result.get("error")
    print(f"[AUTH] Poll result: error={error}, desc={result.get('error_description', '')[:80]}")
    if error == "authorization_pending":
        return {"status": "pending"}
    if error == "expired_token":
        return {"status": "expired"}
    return {"status": "error", "error": error, "description": result.get("error_description", "")}


def fetch_list_data(access_token):
    url = f"https://{SHAREPOINT_HOST}{SHAREPOINT_SITE}/_api/web/lists/getbytitle('{urllib.parse.quote(SHAREPOINT_LIST)}')/items?$top=5000"
    headers = {
        "Accept": "application/json;odata=verbose",
        "Authorization": f"Bearer {access_token}",
    }
    print(f"[DATA] Fetching: {url[:100]}...")
    resp = requests.get(url, headers=headers, timeout=30)
    print(f"[DATA] Response: HTTP {resp.status_code}")
    if resp.status_code != 200:
        raise Exception(f"SharePoint API error: HTTP {resp.status_code} - {resp.text[:300]}")
    data = resp.json()
    results = data.get("d", {}).get("results", data.get("value", []))
    print(f"[DATA] Got {len(results)} items")
    return results


class DashboardHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path.rstrip("/")

        if path == "/" or path == "/dashboard.html":
            self.serve_dashboard()
        elif path == "/api/auth/login":
            self.send_json(start_device_login())
        elif path == "/api/auth/status":
            token = get_token()
            self.send_json({"authenticated": token is not None})
        elif path == "/api/data":
            token = get_token()
            if not token:
                self.send_json({"error": "not_authenticated"}, 401)
                return
            try:
                items = fetch_list_data(token)
                self.send_json({"items": items, "count": len(items)})
            except Exception as e:
                self.send_json({"error": str(e)}, 500)
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/auth/poll":
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length)) if length else {}
            self.send_json(poll_token(body.get("device_code", "")))
        else:
            self.send_json({"error": "not_found"}, 404)

    def serve_dashboard(self):
        try:
            with open("dashboard.html", "rb") as f:
                content = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-cache")
            self.end_headers()
            self.wfile.write(content)
        except FileNotFoundError:
            self.send_error(404, "dashboard.html not found")

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())

    def log_message(self, format, *args):
        msg = format % args
        print(f"[{self.address_string()}] {msg}")


if __name__ == "__main__":
    print(f"""
{'='*60}
  SharePoint Activity Dashboard - Proxy Server
{'='*60}

  Site: {SHAREPOINT_HOST}{SHAREPOINT_SITE}
  List: {SHAREPOINT_LIST}

  Open: http://localhost:{PORT}

  To connect, open the dashboard and click "Connect to SharePoint",
  then follow the device code login instructions.

  Press Ctrl+C to stop the server.
{'='*60}
""")
    server = http.server.HTTPServer(("0.0.0.0", PORT), DashboardHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down...")
        server.shutdown()
