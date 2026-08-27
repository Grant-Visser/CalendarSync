# O365 Calendar Sync

Syncs between one **primary** calendar and any number of **child** calendars across different M365 tenants.

| Direction | What's copied |
|---|---|
| Primary → Child | Privacy-safe "Busy" block (time only, no details) |
| Child → Primary | Full event details (title, body, location) |

## Setup

### 1. Register Azure Apps (one per tenant/account)

For **each** account (primary + every child):

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it e.g. `CalendarSync`
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: leave blank (device code flow)
5. After creation: note the **Application (client) ID** and **Directory (tenant) ID**
6. **API permissions** → **Add** → **Microsoft Graph** → **Delegated** → `Calendars.ReadWrite` → **Add**
7. Click **Grant admin consent** (or ask your admin)
8. **Authentication** → enable **"Allow public client flows"**

### 2. Install dependencies

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 3. Configure

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

### 4. First run (interactive auth — one device-code prompt per account)

```bash
source venv/bin/activate
python3 sync.py
```

Tokens are cached after this; no re-auth needed for ~90 days.

### 5. Schedule with cron

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
