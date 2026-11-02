"""
Account configuration, loaded lazily from PRIMARY_* / CHILD_n_* env vars.

Extracted from sync.py so tools that don't need accounts (e.g. authui.py's
module import) can import sync pieces without triggering sys.exit at import
time. sync.py re-exports these names so existing usage keeps working.
"""

import os
import sys


def load_accounts() -> tuple[dict, list[dict]]:
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