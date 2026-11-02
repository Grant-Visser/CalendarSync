"""
MSAL auth plumbing shared by sync.py and authui.py.

Extracted from sync.py so the web auth UI can reuse the exact same
client IDs, authorities, and SerializableTokenCache files without
triggering sync.py's module-level account loading.
sync.py re-imports these names so the README's `_get_token` example
keeps working unchanged.
"""

import os
from pathlib import Path

from msal import PublicClientApplication

SCOPES = ["Calendars.ReadWrite"]
TOKEN_CACHE_DIR = Path(os.getenv("TOKEN_CACHE_DIR", "~/.calendar_sync_tokens")).expanduser()


def msal_token_cache(cache_file: Path):
    from msal import SerializableTokenCache
    cache = SerializableTokenCache()
    if cache_file.exists():
        cache.deserialize(cache_file.read_text())
    return cache


def _build_msal_app(name: str, cfg: dict):
    TOKEN_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_file = TOKEN_CACHE_DIR / f"{name}.json"
    cache = msal_token_cache(cache_file)
    app = PublicClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant_id']}",
        token_cache=cache,
    )
    return app, cache, cache_file