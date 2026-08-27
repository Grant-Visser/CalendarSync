#!/usr/bin/env python3
"""
O365 calendar sync: one primary calendar, many child calendars.

Primary → Children : privacy-safe "Busy" block (time only, no details)
Child   → Primary  : full event details

Run once interactively to authenticate each account; afterwards safe for cron.
"""

import json
import os
import sys
import hashlib
from datetime import datetime, timezone, timedelta
from pathlib import Path

import requests
from msal import PublicClientApplication
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────

SCOPES = ["Calendars.ReadWrite"]
GRAPH = "https://graph.microsoft.com/v1.0"
SYNC_WINDOW_DAYS = int(os.getenv("SYNC_WINDOW_DAYS", "30"))
STATE_FILE = Path(os.getenv("STATE_FILE", "~/.calendar_sync_state.json")).expanduser()
TOKEN_CACHE_DIR = Path(os.getenv("TOKEN_CACHE_DIR", "~/.calendar_sync_tokens")).expanduser()

# Marker so we never re-sync a copy (extended property)
SYNC_MARKER_NS = "String {d4e28c00-1234-5678-abcd-ef0123456789} Name SyncedFrom"


def _load_accounts() -> tuple[dict, list[dict]]:
    """
    Parse PRIMARY_* and CHILD_n_* env vars.
    Returns (primary_cfg, [child_cfg, ...]).
    Children are discovered by scanning for CHILD_1_CLIENT_ID, CHILD_2_CLIENT_ID, ...
    """
    primary = {
        "name": "primary",
        "client_id": os.environ["PRIMARY_CLIENT_ID"],
        "tenant_id": os.environ["PRIMARY_TENANT_ID"],
        "calendar_id": os.getenv("PRIMARY_CALENDAR_ID", "primary"),
    }

    children = []
    i = 1
    while os.getenv(f"CHILD_{i}_CLIENT_ID"):
        children.append({
            "name": f"child{i}",
            "client_id": os.environ[f"CHILD_{i}_CLIENT_ID"],
            "tenant_id": os.environ[f"CHILD_{i}_TENANT_ID"],
            "calendar_id": os.getenv(f"CHILD_{i}_CALENDAR_ID", "primary"),
        })
        i += 1

    if not children:
        sys.exit("No children configured. Set CHILD_1_CLIENT_ID, CHILD_1_TENANT_ID, etc.")

    return primary, children


PRIMARY_CFG, CHILDREN_CFG = _load_accounts()

# ── Auth ──────────────────────────────────────────────────────────────────────

TOKEN_CACHE_DIR.mkdir(parents=True, exist_ok=True)

def _build_msal_app(name: str, cfg: dict):
    cache_file = TOKEN_CACHE_DIR / f"{name}.json"
    cache = msal_token_cache(cache_file)
    app = PublicClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        token_cache=cache,
    )
    return app, cache, cache_file


def _get_token(name: str, cfg: dict, interactive: bool = False) -> str:
    """
    Return a valid access token. In normal (cron) mode, silent refresh only —
    exits with an error if the cache is missing or the refresh token has expired,
    so cron never hangs waiting for a human.
    Pass interactive=True (via `python3 sync.py --auth`) to do the initial device-code flow.
    """
    app, cache, cache_file = _build_msal_app(name, cfg)

    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result:
        if not interactive:
            sys.exit(
                f"[{name}] Token missing or expired. Run: python3 sync.py --auth\n"
                f"  (cache: {cache_file})"
            )
        print(f"\n[{name}] Open this URL to authenticate:\n")
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        sys.exit(f"[{name}] Auth failed: {result.get('error_description')}")

    cache_file.write_text(cache.serialize())
    return result["access_token"]


def msal_token_cache(cache_file: Path):
    from msal import SerializableTokenCache
    cache = SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text())
    return cache


# ── Graph helpers ─────────────────────────────────────────────────────────────

def graph_get(token: str, url: str, params: dict = None) -> dict:
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params)
    r.raise_for_status()
    return r.json()


def graph_post(token: str, url: str, body: dict) -> dict:
    r = requests.post(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=body)
    r.raise_for_status()
    return r.json()


def graph_patch(token: str, url: str, body: dict):
    r = requests.patch(url, headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=body)
    r.raise_for_status()


def graph_delete(token: str, url: str):
    r = requests.delete(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()


# ── Calendar ops ──────────────────────────────────────────────────────────────

def calendar_url(cfg: dict) -> str:
    cal = cfg["calendar_id"]
    if cal == "primary":
        return f"{GRAPH}/me/calendar"
    return f"{GRAPH}/me/calendars/{cal}"


def list_events(token: str, cfg: dict) -> list:
    """Fetch events in the sync window, page through all results."""
    now = datetime.now(timezone.utc)
    start = (now - timedelta(days=SYNC_WINDOW_DAYS)).isoformat()
    end = (now + timedelta(days=SYNC_WINDOW_DAYS)).isoformat()

    base = calendar_url(cfg)
    url = f"{base}/calendarView"
    params = {
        "$select": "id,subject,body,start,end,location,isAllDay,showAs,isCancelled,singleValueExtendedProperties",
        "$expand": f"singleValueExtendedProperties($filter=id eq '{SYNC_MARKER_NS}')",
        "startDateTime": start,
        "endDateTime": end,
        "$top": "100",
    }

    events = []
    while url:
        data = graph_get(token, url, params)
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None  # nextLink already has params baked in
    return events


def is_synced_copy(event: dict) -> str | None:
    """Return the source event ID if this event is a copy we created, else None."""
    props = event.get("singleValueExtendedProperties") or []
    for p in props:
        if p.get("id") == SYNC_MARKER_NS:
            return p["value"]
    return None


def event_fingerprint(event: dict, full_details: bool) -> str:
    """Stable hash of fields that matter for this copy type."""
    data: dict = {"start": event.get("start"), "end": event.get("end")}
    if full_details:
        data["subject"] = event.get("subject")
        data["location"] = (event.get("location") or {}).get("displayName")
        data["body"] = (event.get("body") or {}).get("content", "")[:500]
    return hashlib.sha1(json.dumps(data, sort_keys=True).encode()).hexdigest()


def _busy_body(event: dict, source_id: str) -> dict:
    """Minimal 'Busy' block for primary→child direction."""
    return {
        "subject": "Busy",
        "start": event["start"],
        "end": event["end"],
        "isAllDay": event.get("isAllDay", False),
        "showAs": "busy",
        "singleValueExtendedProperties": [
            {"id": SYNC_MARKER_NS, "value": source_id}
        ],
    }


def _full_body(event: dict, source_id: str) -> dict:
    """Full event copy for child→primary direction."""
    return {
        "subject": event.get("subject", "(No title)"),
        "start": event["start"],
        "end": event["end"],
        "isAllDay": event.get("isAllDay", False),
        "showAs": event.get("showAs", "busy"),
        "body": event.get("body"),
        "location": event.get("location"),
        "singleValueExtendedProperties": [
            {"id": SYNC_MARKER_NS, "value": source_id}
        ],
    }


def create_copy(token: str, cfg: dict, event: dict, source_id: str, full_details: bool) -> str:
    build = _full_body if full_details else _busy_body
    base = calendar_url(cfg)
    result = graph_post(token, f"{base}/events", build(event, source_id))
    return result["id"]


def update_copy(token: str, event_id: str, event: dict, full_details: bool):
    if full_details:
        body = {
            "subject": event.get("subject", "(No title)"),
            "start": event["start"],
            "end": event["end"],
            "isAllDay": event.get("isAllDay", False),
            "showAs": event.get("showAs", "busy"),
            "body": event.get("body"),
            "location": event.get("location"),
        }
    else:
        body = {
            "subject": "Busy",
            "start": event["start"],
            "end": event["end"],
            "isAllDay": event.get("isAllDay", False),
            "showAs": "busy",
        }
    graph_patch(token, f"{GRAPH}/me/events/{event_id}", body)


def delete_event(token: str, event_id: str):
    try:
        graph_delete(token, f"{GRAPH}/me/events/{event_id}")
    except requests.HTTPError as e:
        if e.response.status_code != 404:
            raise


# ── State ─────────────────────────────────────────────────────────────────────

def load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {}  # {src_id: {"copy_id": str, "fingerprint": str, "direction": "cal1->cal2"|...}}


def save_state(state: dict):
    STATE_FILE.write_text(json.dumps(state, indent=2))


# ── Sync logic ────────────────────────────────────────────────────────────────

def sync_direction(src_name: str, dst_name: str, src_token: str, dst_token: str,
                   src_cfg: dict, dst_cfg: dict, state: dict, full_details: bool):
    direction = f"{src_name}->{dst_name}"
    mode = "full" if full_details else "busy-only"
    print(f"\n── {direction} ({mode}) ──")

    src_events = list_events(src_token, src_cfg)
    dst_events = list_events(dst_token, dst_cfg)

    # Build lookup: source_id -> copy in destination
    dst_copies = {is_synced_copy(e): e for e in dst_events if is_synced_copy(e)}

    # Originals only — skip copies we received from elsewhere
    src_originals = {e["id"]: e for e in src_events if not is_synced_copy(e)}

    created = updated = deleted = skipped = 0

    for src_id, src_event in src_originals.items():
        if src_event.get("isCancelled"):
            continue

        fp = event_fingerprint(src_event, full_details)
        state_key = f"{direction}:{src_id}"
        existing = state.get(state_key)

        if src_id in dst_copies:
            dst_event = dst_copies[src_id]
            if existing and existing.get("fingerprint") == fp:
                skipped += 1
                continue
            update_copy(dst_token, dst_event["id"], src_event, full_details)
            state[state_key] = {"copy_id": dst_event["id"], "fingerprint": fp}
            updated += 1
        else:
            if existing:
                del state[state_key]
            copy_id = create_copy(dst_token, dst_cfg, src_event, src_id, full_details)
            state[state_key] = {"copy_id": copy_id, "fingerprint": fp}
            created += 1

    # Delete copies whose source has left the sync window
    for key in [k for k in state if k.startswith(f"{direction}:")]:
        src_id = key.split(":", 2)[2]
        if src_id not in src_originals:
            delete_event(dst_token, state[key]["copy_id"])
            del state[key]
            deleted += 1

    print(f"  created={created} updated={updated} deleted={deleted} skipped={skipped}")


def main():
    interactive = "--auth" in sys.argv
    all_cfgs = [PRIMARY_CFG] + CHILDREN_CFG

    if interactive:
        print("Auth mode — you will be prompted to sign in for each account.\n")

    tokens = {cfg["name"]: _get_token(cfg["name"], cfg, interactive) for cfg in all_cfgs}

    if interactive:
        print("\nAll accounts authenticated. You can now run the script without --auth.")
        return
    state = load_state()

    primary_name = PRIMARY_CFG["name"]
    primary_token = tokens[primary_name]

    for child in CHILDREN_CFG:
        child_name = child["name"]
        child_token = tokens[child_name]

        # Primary → Child: busy blocks only
        sync_direction(primary_name, child_name, primary_token, child_token,
                       PRIMARY_CFG, child, state, full_details=False)

        # Child → Primary: full details
        sync_direction(child_name, primary_name, child_token, primary_token,
                       child, PRIMARY_CFG, state, full_details=True)

    save_state(state)
    print("\nDone.")


if __name__ == "__main__":
    main()
