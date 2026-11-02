#!/usr/bin/env python3
"""
Local web auth UI for CalendarSync (Option B — keeper.sh-style).

Shows a per-account token-health card and a "Sign in / Re-auth" button.
Uses the same MSAL PublicClientApplications and SerializableTokenCache
files as sync.py, so a successful web sign-in immediately unblocks cron.

Run:  venv/bin/python3 authui.py   (port from .env WEBAUTH_PORT, default 54231)

Requires one-time manual Azure steps: add BOTH redirect URIs
  http://localhost:54231/auth/callback   (direct access)
  https://calendarsync-auth.home.orangelabs.dev/auth/callback   (via NPM proxy)
to each app registration (keep http://localhost so `sync.py --auth` works).
The redirect URI is built per-request from Host / X-Forwarded-Proto headers.
"""

import json
import html
import time
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from dotenv import load_dotenv
import msal

from accounts import load_accounts
from auth import SCOPES, TOKEN_CACHE_DIR, _build_msal_app

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

WEBAUTH_PORT = int(__import__("os").getenv("WEBAUTH_PORT", "54231"))
CLAIMS_FILE = Path("~/.calendar_sync_claims.json").expanduser()

# In-flight auth-code flows and flash messages: single-user local app, so a
# process-local dict keyed by state (with expiry) is sufficient.
FLOWS: dict[str, tuple[dict, str, float]] = {}  # state -> (flow, name, expires_at)
FLOW_TTL = 600  # seconds

_flash: dict[str, tuple[str, str, float]] = {}  # id -> (kind, text, expires_at)
_flash_seq = 0


def flash_set(kind: str, text: str) -> str:
    global _flash_seq
    _flash_seq += 1
    key = str(_flash_seq)
    _flash[key] = (kind, text, time.time() + 60)
    return key


def flash_pop(key: str) -> tuple[str, str] | None:
    entry = _flash.pop(key, None)
    if entry and entry[2] > time.time():
        return entry[0], entry[1]
    return None


def gc():
    now = time.time()
    for d in (FLOWS, _flash):
        for k in [k for k, v in d.items() if v[-1] < now]:
            del d[k]


# ── Claims-file hook (Option A prerequisite) ─────────────────────────────────

def pending_claims(name: str) -> str | None:
    """
    Read a persisted claims challenge for `name`, if any.

    Expected file format (written by Option A's 401 handler, not implemented yet):
        {"primary": {"claims": "<url-encoded claims>", "timestamp": 169...}, ...}
    """
    if not CLAIMS_FILE.exists():
        return None
    try:
        data = json.loads(CLAIMS_FILE.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    entry = data.get(name)
    if isinstance(entry, dict) and entry.get("claims"):
        return entry["claims"]
    return None


def claims_challenge_pending(name: str) -> bool:
    return pending_claims(name) is not None


# ── Token health ──────────────────────────────────────────────────────────────

def token_health(name: str, cfg: dict) -> dict:
    """
    Classify an account without any interactive step:
      - "no account"      — cache empty / wrong tenant
      - "claims pending"  — 401 challenge persisted, re-auth will satisfy it
      - "silent ok"       — acquire_token_silent succeeded
      - "needs re-auth"   — account exists but silent refresh fails
    """
    app, cache, cache_file = _build_msal_app(name, cfg)
    accounts = app.get_accounts()
    if not accounts:
        return {"state": "no-account", "detail": "no account in cache",
                "cache": str(cache_file)}

    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    username = accounts[0].get("username") or accounts[0].get("home_account_id") or "?"

    if result and "access_token" in result:
        # Silent acquisition may itself return a claims challenge dict.
        if result.get("error") == "interaction_required" or "claims" in result:
            state, detail = "claims", "claims challenge pending"
        else:
            state, detail = "ok", "silent token OK"
    elif result and result.get("error") == "interaction_required" and result.get("claims"):
        state, detail = "claims", "claims challenge pending"
    else:
        state, detail = "expired", "needs re-auth"

    if claims_challenge_pending(name):
        state, detail = "claims", "claims challenge pending (re-auth will satisfy it)"

    return {"state": state, "detail": detail, "user": username,
            "cache": str(cache_file)}


# ── HTML ──────────────────────────────────────────────────────────────────────

PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>CalendarSync auth</title>
<script src="https://unpkg.com/htmx.org@1.9.12"></script>
<style>
 body {{ font-family: system-ui, sans-serif; max-width: 720px; margin: 2rem auto; padding: 0 1rem; }}
 .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
 .ok {{ border-left: 6px solid #2e7d32; }}
 .expired {{ border-left: 6px solid #f9a825; }}
 .claims {{ border-left: 6px solid #ef6c00; }}
 .no-account {{ border-left: 6px solid #c62828; }}
 .btn {{ display: inline-block; padding: .5rem 1rem; background: #1a73e8; color: #fff;
         border-radius: 6px; text-decoration: none; border: none; cursor: pointer; }}
 .flash {{ padding: .6rem 1rem; border-radius: 6px; margin: 1rem 0; }}
 .flash.ok {{ background: #e8f5e9; }} .flash.err {{ background: #ffebee; }}
 small {{ color: #666; }}
</style>
</head>
<body>
<h1>CalendarSync — account auth</h1>
<div id="flash">{flash}</div>
<div id="cards" hx-get="/cards" hx-trigger="load delay:0s, every 60s" hx-swap="innerHTML"></div>
<p><small>Tokens are written to the same cache files sync.py uses, so cron keeps
working right after a web sign-in. Azure redirect URIs required (exact match):
<code>http://localhost:{port}/auth/callback</code> and
<code>https://calendarsync-auth.home.orangelabs.dev/auth/callback</code>.
Single-user LAN-only tool — never expose publicly.</small></p>
</body>
</html>"""

LABELS = {"ok": "Silent OK", "expired": "Needs re-auth",
          "claims": "Claims challenge pending", "no-account": "Not signed in"}


def render_cards() -> str:
    gc()
    primary, children = load_accounts()
    out = []
    for cfg in [primary] + children:
        h = token_health(cfg["name"], cfg)
        cls = h["state"].replace("no-account", "no-account")
        sign_in = ("" if h["state"] == "ok"
                   else f'<a class="btn" href="/signin/{html.escape(cfg["name"])}">'
                        f'Sign in / Re-auth</a>')
        out.append(
            f'<div class="card {cls}">'
            f'<h3>{html.escape(cfg["name"])} <small>{html.escape(h.get("user", ""))}</small></h3>'
            f'<p><b>{LABELS.get(h["state"], h["state"])}</b> — {html.escape(h["detail"])}</p>'
            f'{sign_in}'
            f'<p><small>cache: {html.escape(h["cache"])}</small></p>'
            f'</div>')
    return "\n".join(out)


# ── HTTP handler ──────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):

    def _redirect(self, location: str):
        self.send_response(302)
        self.send_header("Location", location)
        self.end_headers()

    def _html(self, body: str, code: int = 200):
        data = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):  # noqa: N802 — http.server API
        parsed = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed.query)
        gc()

        if parsed.path == "/":
            fk = (qs.get("flash") or [None])[0]
            kind, text = flash_pop(fk) if fk else (None, None)
            flash_html = (f'<div class="flash {kind}">{html.escape(text)}</div>'
                          if kind else "")
            self._html(PAGE.format(flash=flash_html, port=WEBAUTH_PORT))

        elif parsed.path == "/cards":
            self._html(render_cards())

        elif parsed.path.startswith("/signin/"):
            name = parsed.path[len("/signin/"):]
            cfg = self._cfg_for(name)
            if not cfg:
                self._redirect("/?flash=" + flash_set("err", f"Unknown account {name!r}"))
                return
            app, _, _ = _build_msal_app(name, cfg)
            flow = app.initiate_auth_code_flow(
                SCOPES, redirect_uri=self._redirect_uri())
            FLOWS[flow["state"]] = (flow, name, time.time() + FLOW_TTL)
            self._redirect(flow["auth_uri"])

        elif parsed.path == "/auth/callback":
            if "error" in qs:
                desc = (qs.get("error_description") or [qs["error"]])[0]
                self._redirect("/?flash=" + flash_set("err", f"Sign-in failed: {desc}"))
                return
            state = (qs.get("state") or [None])[0]
            entry = FLOWS.pop(state, None) if state else None
            if not entry:
                self._redirect("/?flash=" + flash_set("err", "Unknown or expired sign-in flow"))
                return
            flow, name, _ = entry
            cfg = self._cfg_for(name)
            app, cache, cache_file = _build_msal_app(name, cfg)
            result = app.acquire_token_by_auth_code_flow(
                flow, qs, claims_challenge=pending_claims(name))
            if "access_token" in result:
                TOKEN_CACHE_DIR.mkdir(parents=True, exist_ok=True)
                cache_file.write_text(cache.serialize())
                self._redirect("/?flash=" + flash_set("ok", f"{name}: signed in as "
                                                      f"{result.get('id_token_claims', {}).get('preferred_username', '?')}"))
            else:
                err = result.get("error_description") or result.get("error")
                self._redirect("/?flash=" + flash_set("err", f"{name}: auth failed — {err}"))

        else:
            self._html("<h1>404</h1>", code=404)

    def _redirect_uri(self) -> str:
        """
        Build the OAuth redirect URI from the incoming request so the UI works
        both directly (http://localhost:54231) and behind a reverse proxy
        (https://calendarsync-auth.home.orangelabs.dev).

        Azure matches redirect URIs as EXACT strings — the same path is used
        for every scheme/host, only scheme+host vary. Scheme comes from
        X-Forwarded-Proto (first hop) when proxied, else plain http.
        """
        proto = self.headers.get("X-Forwarded-Proto")
        scheme = proto.split(",")[0].strip() if proto else "http"
        host = self.headers.get("Host") or f"localhost:{WEBAUTH_PORT}"
        return f"{scheme}://{host}/auth/callback"

    def _cfg_for(self, name: str) -> dict | None:
        primary, children = load_accounts()
        for cfg in [primary] + children:
            if cfg["name"] == name:
                return cfg
        return None

    def log_message(self, fmt, *args):  # quieter logs
        print(f"[authui] {self.address_string()} {fmt % args}")


def main():
    primary, children = load_accounts()  # fail fast on missing env
    names = [c["name"] for c in [primary] + children]
    print(f"CalendarSync auth UI — accounts: {', '.join(names)}")
    print(f"Listening on http://0.0.0.0:{WEBAUTH_PORT} (direct: http://localhost:{WEBAUTH_PORT}; "
          f"or via reverse proxy, e.g. https://calendarsync-auth.home.orangelabs.dev)")
    print("LAN-only tool — do not expose publicly.")
    server = ThreadingHTTPServer(("0.0.0.0", WEBAUTH_PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
