#!/usr/bin/env python3
"""
Streamlined setup: creates Azure app registrations for CalendarSync
using the Azure CLI (az), then writes .env for you.

Requirements: Azure CLI installed and accessible as `az`.
Run once per machine. Re-running is safe — skips existing apps.

Usage:
    python3 setup.py
"""

import json
import os
import subprocess
import sys
from pathlib import Path


def az(*args, check=True) -> dict | str | None:
    """Run an az command, return parsed JSON or raw stdout."""
    cmd = ["az", *args]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if check and result.returncode != 0:
        sys.exit(f"az error: {result.stderr.strip()}")
    if not result.stdout.strip():
        return None
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        return result.stdout.strip()


def az_ok() -> bool:
    result = subprocess.run(["az", "--version"], capture_output=True)
    return result.returncode == 0


def current_tenant() -> str:
    account = az("account", "show")
    return account["tenantId"]


def login(label: str) -> str:
    """Interactive browser login; returns tenant ID."""
    print(f"\n[{label}] A browser window will open — sign in with the account for this calendar.")
    az("login", "--allow-no-subscriptions")
    tid = current_tenant()
    print(f"  ✓ Signed in to tenant: {tid}")
    return tid


def get_or_create_app(display_name: str) -> tuple[str, str]:
    """
    Return (client_id, tenant_id) for a new or existing app registration.
    Skips creation if an app with this name already exists.
    """
    existing = az("ad", "app", "list", "--display-name", display_name, "--query", "[0]")
    if existing:
        client_id = existing["appId"]
        print(f"  ↩ App '{display_name}' already exists: {client_id}")
    else:
        app = az(
            "ad", "app", "create",
            "--display-name", display_name,
            "--sign-in-audience", "AzureADMyOrg",
            "--public-client-redirect-uris", "https://login.microsoftonline.com/common/oauth2/nativeclient",
        )
        client_id = app["appId"]
        print(f"  ✓ Created app '{display_name}': {client_id}")

        # Add Calendars.ReadWrite delegated permission
        # Resource: Microsoft Graph (00000003-0000-0000-c000-000000000000)
        # Scope ID for Calendars.ReadWrite: 465a38f9-76ea-45b9-9f34-9e8b0d4b0b42  (well-known)
        az(
            "ad", "app", "permission", "add",
            "--id", client_id,
            "--api", "00000003-0000-0000-c000-000000000000",
            "--api-permissions", "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42=Scope",
        )

        # Enable public client flows (required for device-code auth)
        az(
            "ad", "app", "update",
            "--id", client_id,
            "--enable-mobile-and-desktop-flows", "true",
        )

        print(f"  ✓ Calendars.ReadWrite permission added, public client flows enabled")

    tenant_id = current_tenant()
    return client_id, tenant_id


def write_env(entries: list[tuple[str, str, str]]) -> Path:
    """
    entries: [(prefix, client_id, tenant_id), ...]
    Writes .env, preserving any existing non-CalendarSync lines.
    """
    env_path = Path(".env")
    existing_lines = []
    managed_prefixes = {"PRIMARY_", "CHILD_"}

    if env_path.exists():
        for line in env_path.read_text().splitlines():
            key = line.split("=")[0].strip()
            if not any(key.startswith(p) for p in managed_prefixes):
                existing_lines.append(line)

    new_lines = []
    for prefix, client_id, tenant_id in entries:
        new_lines += [
            f"{prefix}CLIENT_ID={client_id}",
            f"{prefix}TENANT_ID={tenant_id}",
            f"{prefix}CALENDAR_ID=primary",
        ]

    output = "\n".join(existing_lines + [""] + new_lines + [""]).lstrip("\n")
    env_path.write_text(output)
    return env_path


def main():
    if not az_ok():
        sys.exit(
            "Azure CLI not found. Install it:\n"
            "  https://docs.microsoft.com/cli/azure/install-azure-cli\n"
            "  Ubuntu: curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash"
        )

    print("CalendarSync Setup")
    print("=" * 40)
    print("You'll be asked to sign in once per calendar account.")
    print("Each sign-in opens a browser window.\n")

    n_children = 0
    while True:
        try:
            n_children = int(input("How many child calendars do you want to set up? "))
            if n_children >= 1:
                break
            print("Need at least 1.")
        except ValueError:
            pass

    entries = []  # (env_prefix, client_id, tenant_id)

    # Primary
    print("\n── Primary calendar ──")
    login("primary")
    client_id, tenant_id = get_or_create_app("CalendarSync-Primary")
    entries.append(("PRIMARY_", client_id, tenant_id))

    # Children
    for i in range(1, n_children + 1):
        print(f"\n── Child calendar {i} ──")
        login(f"child{i}")
        client_id, tenant_id = get_or_create_app(f"CalendarSync-Child{i}")
        entries.append((f"CHILD_{i}_", client_id, tenant_id))

    env_path = write_env(entries)
    print(f"\n{'=' * 40}")
    print(f"✓ Written to {env_path.resolve()}")
    print("\nNext steps:")
    print("  1. python3 sync.py --auth   ← consent prompts, once per account")
    print("  2. python3 sync.py          ← test a manual sync")
    print("  3. Add to cron (see README)")


if __name__ == "__main__":
    main()
