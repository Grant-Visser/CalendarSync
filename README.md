# O365 Calendar Sync

Syncs between one **primary** calendar and any number of **child** calendars across different M365 tenants.

| Direction | What's copied |
|---|---|
| Primary → Child | Privacy-safe "Busy" block (time only, no details) |
| Child → Primary | Full event details (title, body, location) |

## Setup

### Prerequisites

- Python 3.10+
- Azure CLI (required for `setup.py`'s app registration step; not pip-installable — use the [official installer](https://learn.microsoft.com/cli/azure/install-azure-cli) for your OS): `curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash`

### Streamlined setup (recommended)

`setup.py` creates the Azure app registrations for you — no portal clicking.

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

python3 setup.py
```

It will ask how many child calendars you have, then open a browser sign-in for each account and write your `.env` automatically. After it finishes:

```bash
python3 sync.py --auth   # one-time consent prompt per account
python3 sync.py          # verify it works
```

Then add to cron (see [Scheduling](#scheduling)).

---

### Manual setup (alternative)

If you prefer to create app registrations yourself in the Azure portal:

#### 1. Register Azure Apps (one per tenant/account)

For **each** account (primary + every child):

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it e.g. `CalendarSync`
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: leave blank (device code flow)
5. After creation: note the **Application (client) ID** and **Directory (tenant) ID**
6. **API permissions** → **Add** → **Microsoft Graph** → **Delegated** → `Calendars.ReadWrite` → **Add**
7. **Authentication** → enable **"Allow public client flows"**
   > You don't need to click "Grant admin consent". When you run `--auth`, Microsoft will show a standard consent prompt for your own account — just click Accept. Admin consent is only needed if your IT department has disabled per-user consent for the entire tenant (you'll see an `AADSTS65001` error if so).

#### 2. Install dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

#### 3. Configure

```bash
cp .env.example .env
# Edit .env — set PRIMARY_* and one CHILD_n_* block per child account
```

To add more children, add `CHILD_3_*`, `CHILD_4_*`, etc. — the script discovers them automatically.

To find a non-default calendar ID (if you want a specific sub-calendar):
```bash
source venv/bin/activate
python3 -c "
import os; from dotenv import load_dotenv; load_dotenv()
from sync import _get_token, PRIMARY_CFG, GRAPH, graph_get
t = _get_token('primary', PRIMARY_CFG)
cals = graph_get(t, f'{GRAPH}/me/calendars', {'\$select': 'id,name'})
for c in cals['value']: print(c['id'], c['name'])
"
```

#### 4. Authenticate

```bash
python3 sync.py --auth   # device-code prompts, once per account
```

---

## Web auth UI (recommended for re-auth)

A tiny local web app (`authui.py`, stdlib `http.server` + htmx) that shows a
per-account token-health card and a **Sign in / Re-auth** button — no SSH +
`--auth` dance when a Conditional Access policy expires your refresh token.

```bash
source venv/bin/activate
python3 authui.py          # binds 0.0.0.0:$WEBAUTH_PORT (default 54231)
```

Then open either `http://localhost:54231` (direct) or
`https://calendarsync-auth.orangelabs.home` (via Nginx Proxy Manager;
LAN/WireGuard only). This is a **single-user, LAN-only tool — never expose it
publicly**. The OAuth redirect URI is built from the incoming request's Host
and `X-Forwarded-Proto` headers, so both entry points produce a redirect URI
Azure will accept (exact string match — register both URIs, see below).

Each card shows one of:

- **Silent OK** — token cache is healthy, cron sync will work
- **Needs re-auth** — click the button to sign in again
- **Claims challenge pending** — a Graph 401 claims challenge was persisted
  (`~/.calendar_sync_claims.json`); signing in via the button satisfies it
- **Not signed in** — no account in the cache yet

A successful web sign-in writes the same `SerializableTokenCache` files
`sync.py` uses (`~/.calendar_sync_tokens/<name>.json`), so the next cron run
works immediately. The cron path (`sync.py` without `--auth`) is untouched.

### Manual Azure portal step (required, once per app)

The app registrations already exist and there is no Azure CLI/portal access
from this machine, so add the redirect URI by hand in the
[Azure portal](https://portal.azure.com)
(**Entra ID / Azure Active Directory → App registrations → <your app> → Authentication**):

1. **Mobile and desktop applications** platform: add
   `http://localhost:54231/auth/callback`
   (the port must match `WEBAUTH_PORT` in `.env` exactly — redirect URIs are
   string-matched, no wildcards)
2. **Web** platform: add
   `https://calendarsync-auth.orangelabs.home/auth/callback`
   (served through Nginx Proxy Manager; LAN/WireGuard only — never expose
   this publicly. Keep the app **public**: no client secret, PKCE applies.)
3. Keep the existing desktop `http://localhost` redirect so `sync.py --auth`
   keeps working — URIs are additive.
4. Ensure **"Allow public client flows"** stays set to **Yes**
   (`isFallbackPublicClient=true` in the manifest).

Do this for **each** app registration (primary + every child).

> Note: the web UI uses auth-code + PKCE (`initiate_auth_code_flow`), so no
> client secret is needed; the apps remain public clients.

### Running it long-term

The UI is only needed when auth is required, so it doesn't have to run all
the time. If you want it always available on the host (or reachable at the
NPM hostname), a small systemd user unit or `tmux` session works; cron
`sync.py` runs are unaffected by the UI being up or down (they never touch
the HTTP port).

### Scheduling

#### Docker (recommended for servers)

Auth must happen on the host first — the container has no browser:

```bash
# 1. Authenticate on the host (opens browser per account)
python3 sync.py --auth

# 2. Start the container — it reads tokens and state directly from the host
docker compose up -d
docker compose logs -f
```

The `.env` file is passed in via `env_file`, and the token cache and state file are bind-mounted from their host paths (defaulting to `~/.calendar_sync_tokens` and `~/.calendar_sync_state.json`, or whatever you set in `.env`).

If tokens ever expire, re-run `python3 sync.py --auth` on the host then `docker compose restart`.

#### Cron (alternative)

```bash
crontab -e
```

Add (every 15 minutes):
```
*/15 * * * * cd /path/to/CalendarSync && venv/bin/python3 sync.py >> /var/log/calendar_sync.log 2>&1
```

## How it works

- Syncs events within ±`SYNC_WINDOW_DAYS` days of now
- Each copy is tagged with a hidden extended property — the script never double-syncs a copy
- State file (`~/.calendar_sync_state.json`) tracks copy IDs and content fingerprints to skip unchanged events
- Deleted/cancelled source events cause the copy to be deleted in the destination
