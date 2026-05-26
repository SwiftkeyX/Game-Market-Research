---
name: genre-viability-data
description: >
  Usage: (no arg) view ratings | [update] refresh from live data | [add][genre] add new genre |
  [remove][genre] delete a genre. Data lives in Google Sheets (browser-viewable). Warns when data
  is older than 1 month. History versioned as tabs in the same sheet.
---

# Steam Genre Viability Reference

**Storage:** Google Sheets (open in any browser — no license required)
**Last updated:** May 24, 2026
**Next refresh due:** June 24, 2026

> Hit rate = % of released games reaching 1000+ reviews ≈ $150K revenue

---

## Arguments

```
/genre-viability-data                    # read sheet and display current ratings
/genre-viability-data [update]           # fetch live data, update ratings, save history tab
/genre-viability-data [add][genre name]  # research new genre and append row to sheet
/genre-viability-data [remove][genre]    # delete a genre row from the sheet
```

**Staleness check (run on every invocation):**
Compare today's date against **Next refresh due** above. If overdue:
> ⚠️ Data is overdue for a refresh. Run `/genre-viability-data [update]` to update.

Still display current ratings below the warning.

---

## Config

- **`CREDENTIALS_FILE`** — `C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json`
- **`SHEET_URL`** — `https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing`

---

## First-Time Setup (run once)

If Python or gspread is not yet installed, guide the user through these steps:

```
1. Install Python:
   winget install Python.Python.3.12

2. Install libraries:
   pip install gspread google-auth

3. Create a new Google Sheet in your browser.

4. Set up a service account:
   - Go to console.cloud.google.com
   - Create a project → Enable Google Sheets API
   - IAM & Admin → Service Accounts → Create service account
   - Keys → Add Key → JSON → download the file

5. Share your Google Sheet with the service account email
   (found in the JSON file as "client_email") — give it Editor role.

6. Tell Claude:
   - The full path to the JSON key file
   - The URL of your Google Sheet
```

After setup, run the migration to push the 30 existing genres to the sheet:

```python
# migrate.py — run once to populate the sheet
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

HEADERS = ["Verdict","Genre","Solo Target","Team of 4","Hit Rate","Trend","Notes"]

GENRES = [
  ["GO","Roguelike Deckbuilder","$100K-$500K","$500K-$2M","5.1%","Growing","3 hits Q1 2026 (normally 1/yr). Best solo genre. 200-500 cards manageable."],
  ["GO","Idle / Incremental","$100K-$300K","$200K-$600K","~3% growing","Golden age","3 hits 2022 -> 16 in 2024 -> 12 in Q2 2025. 2-3 month solo build. Low art bar."],
  ["GO","Job Simulation 3D","$150K-$400K","$500K-$2M","4.1%","Overtaking Mgmt","Blue-collar 3D first-person dominates. Hard solo without 3D skills."],
  ["GO","Horror (any combo)","$100K-$400K","$300K-$1.5M","3.2%","Cooling Q1 2026","Flexible: Horror+Casino, Horror+Idle, Horror+Co-op all work."],
  ["GO","Friend-Slop Co-op","$100K-$500K","$300K-$1.5M","Top Q1 2026","Rising","Most hits Q1 2026. Build in 4-8 weeks. Viral streaming potential."],
  ["GO","Narrative Story-Driven","$100K-$500K","$300K-$1.5M","51 hits 2025","Strongly rising","#1 genre by hit count 2025. Story + light gameplay systems."],
  ["GO","Pixel Art / Indie RPG","$100K-$400K","$300K-$1.5M","~2.4% stable","Consistent","18 hits Q2 2025. Top-down, 8-15hrs focused scope."],
  ["GO","Colony Sim / Factory","N/A (solo skip)","$500K-$3M","6.4% hybrid","Rising","48% of top Next Fest demos. Team only. Expect 2-4yr dev time."],
  ["GO","Farming / Life Sim","$100K-$500K","$300K-$1.5M","Consistent hits","Stable","Multiple hits annually. Need hook beyond just farming."],
  ["GO","Auto-battler","$100K-$400K","$300K-$1.5M","Phase 3 emerging","Emerging","Low saturation window open. Low art burden, high systems depth."],
  ["CAUTION","Tower Defense Hybrid","$100K-$400K","$300K-$1.5M","Reviving","Rising (needs meta)","5 hits Q2 2025 vs 8 all 2024. Meta layer non-negotiable."],
  ["CAUTION","Tactical Strategy RPG","$100K-$500K","$400K-$2M","~3-4%","Steady","Design-heavy not art-heavy. Scope balloons fast; strict discipline needed."],
  ["CAUTION","Open World Survival Craft","N/A (solo skip)","$500K-$3M","20.8%","Dropped from top 5","Highest hit rate but scope brutal (2-4 years). Team only."],
  ["CAUTION","Action Roguelite","$150K-$500K","$500K-$2M","~2-3%","Stable/crowded","Dead Cells + Hades set punishing quality bar. Hard solo without art skills."],
  ["CAUTION","Cozy Management","$100K-$400K","$300K-$1.2M","Moderate","Stable","Management-cozy hits; adventure-cozy does not. Build loop first."],
  ["CAUTION","Metroidvania","$150K-$600K","$400K-$2M","~4% / 0 Q1 hits","Declining freq.","Zero hits Q1 2025. Hollow Knight bar brutal. Must be great, not just good."],
  ["CAUTION","Visual Novel (pure)","$50K-$200K","$100K-$500K","Skewed (Asian pub)","Growing (skewed)","51 hits mostly Chinese FMV. Western pure VN rate far lower than aggregate."],
  ["CAUTION","JRPG","$100K-$400K","$300K-$1.2M","Gatekept","Stable","Dominated by Atlus/Falcom back-catalog. Budget 2-3 years for modest scope."],
  ["CAUTION","3D Platformer","$80K-$300K","$200K-$800K","1.46%","No hits Q1 2025","3D physics/camera exponentially harder. AAA polish expected."],
  ["CAUTION","Rage / Precision Platformer","$50K-$200K","$100K-$400K","Phase 4","Declining","Window closing since 2022. Viral streaming hook essential."],
  ["AVOID","2D Platformer","N/A","N/A","0.18%","Zero hits Q1 2026","Near-zero. Free browser games dominate."],
  ["AVOID","Point-and-Click Adventure","N/A","N/A","0.18%","Declining","LucasArts quality expected. Near-zero modern indie revenue."],
  ["AVOID","Multiplayer Shooter","N/A","N/A","Dies post-launch","Collapses fast","Steel Hunters: 4473 -> 45 players. Cannot sustain playerbase."],
  ["AVOID","Pure Puzzle","N/A","N/A","0.34%","No improvement","As primary tag fails. Players get free content on mobile."],
  ["AVOID","Bullet Heaven Clone","N/A","N/A","<1%","Saturated 2023","Only 1 hit in all of 2024. Window closed."],
  ["AVOID","VR Game","N/A","N/A","0 hits Q1 2026","Declining","VR has no audience at scale. Hardware penetration too low."],
  ["AVOID","4X Strategy","N/A","N/A","Near zero","No improvement","Civ/Stellaris dominate. 3-5yr dev minimum."],
  ["AVOID","Racing Game","N/A","N/A","Near zero","AAA dominated","Forza/GT capture all intent. Physics judged harshly."],
  ["AVOID","Fighting Game","N/A","N/A","Near zero","Gatekept","Capcom/ArcSys dominate. Balancing + rollback netcode required."],
  ["AVOID","Music / Rhythm","N/A","N/A","Near zero paid","F2P kills it","Osu free, Muse Dash F2P. No paid market without compelling IP hook."],
]

# Write Current Ratings tab
try:
    ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
    ws.clear()
except gspread.WorksheetNotFound:
    ws = sh.add_worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)", rows=100, cols=10)
ws.append_row(HEADERS)
ws.append_rows(GENRES)

# Write v1 history snapshot tab
try:
    hist = sh.worksheet("v1 - May 2026")
    hist.clear()
except gspread.WorksheetNotFound:
    hist = sh.add_worksheet("v1 - May 2026", rows=100, cols=10)
hist.append_row(["SNAPSHOT v1 - May 2026 | READ ONLY | Tier movements: None (first version)"])
hist.append_row(HEADERS)
hist.append_rows(GENRES)

print(f"Done. Written {len(GENRES)} genres to 'Genre Viability Ratings (GO / CAUTION / AVOID)' and 'v1 - May 2026'.")
```

---

## Reading the Sheet

On every invocation, read the sheet with Python. Use the full interpreter path:
**`C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe`**

```python
import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
ws     = client.open_by_url(SHEET_URL).worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
rows   = ws.get_all_records()  # returns list of dicts keyed by header row
```

Display results grouped by Verdict: GO first, then CAUTION, then AVOID.

## Color Scheme

Applied to both the Google Sheet and any output displays:

| Verdict | Background | Hex |
|---|---|---|
| Header row | Blue | `#CFE2F3` → RGB(0.812, 0.886, 0.953) |
| GO | Green | `#D9EAD3` → RGB(0.851, 0.918, 0.827) |
| CAUTION | Yellow | `#FFF2CC` → RGB(1.0, 0.949, 0.800) |
| AVOID | Red | `#FFE6E6` → RGB(1.0, 0.902, 0.902) |

When re-applying formatting after any write operation (update/add/remove), use this pattern:

```python
def apply_formatting(sh, ws):
    rows = ws.get_all_records()
    sheet_id = ws.id
    COLOR_HEADER  = {"red": 0.812, "green": 0.886, "blue": 0.953}
    COLOR_GO      = {"red": 0.851, "green": 0.918, "blue": 0.827}
    COLOR_CAUTION = {"red": 1.0,   "green": 0.949, "blue": 0.800}
    COLOR_AVOID   = {"red": 1.0,   "green": 0.902, "blue": 0.902}

    ws.format("A1:G1", {"backgroundColor": COLOR_HEADER, "textFormat": {"bold": True, "fontSize": 11}, "horizontalAlignment": "CENTER"})

    for i, row in enumerate(rows, start=2):
        color = {"GO": COLOR_GO, "CAUTION": COLOR_CAUTION, "AVOID": COLOR_AVOID}.get(row["Verdict"], {})
        ws.format(f"A{i}:G{i}", {"backgroundColor": color})
        ws.format(f"A{i}", {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"})
```

---

## [update] Workflow — Refreshing All Ratings

1. Read all rows from "Current Ratings".
2. Search for the latest howtomarketagame.com report:
   ```
   web_search: site:howtomarketagame.com games selling 2026
   ```
3. Fetch the article; extract hit counts and any GO/AVOID signals per genre.
4. For each row, decide: keep / upgrade / downgrade / remove.
5. **Before writing** — save a history snapshot:
   ```python
   # duplicate the sheet to a new versioned tab
   ws_id   = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)").id
   sh.duplicate_sheet(ws_id, new_sheet_name=f"v{N} - {month_year}")
   hist_ws = sh.worksheet(f"v{N} - {month_year}")
   hist_ws.update("A1", [[f"SNAPSHOT v{N} - {month_year} | READ ONLY | Tier movements: {summary}"]])
   ```
6. Apply changes using `ws.update_cell(row, col, value)` for each changed cell — never clear the whole sheet.
7. Update **Last updated** and **Next refresh due** (+1 month) in this SKILL.md file.
8. **Git: export CSV + commit + push**

```python
# --- Git: export CSV + commit + push ---
import csv, subprocess
from pathlib import Path
from datetime import date

BASE = Path(r"C:\Organized Files\My Game Asset\Game-Research")

ratings_ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
csv_path = BASE / "data" / "genre-viability.csv"
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ratings_ws.get_all_values())
print(f"Exported: data/genre-viability.csv")

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git("add", "data/genre-viability.csv")
_git("commit", "-m", f"update: genre viability ratings - {date.today().isoformat()}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
```

**Rating change signals:**
- 🟡→🟢 Upgrade: hit rate rising 2+ quarters, new breakout hits, low saturation
- 🟢→🟡 Downgrade: hit count dropping quarter-over-quarter, major studios entering
- 🟡→🔴 Avoid: hit rate below 1%, no hits for a full quarter, "dead/avoid" language in sources
- 🔴→🟡 Revival: unexpected breakout hit, new mechanical twist proven to work

---

## [add] Workflow — Adding a New Genre

1. **Check it isn't already tracked:**
   ```python
   rows = ws.get_all_records()
   if any(r["Genre"].lower() == genre.lower() for r in rows):
       # tell user it already exists and show current entry
   ```
2. **Research** using live web search:
   ```
   web_search: site:howtomarketagame.com [genre] 2026
   web_search: games-stats.com steam tag=[genre] revenue hit rate
   web_search: indie [genre] steam revenue 2025 2026
   ```
3. Collect: hit rate, trend, Solo target, Team of 4 target, 2–3 example games, saturation signals.
4. Assign a verdict using the rating change signals above.
5. **Append atomically** (safe — does not touch existing rows):
   ```python
   ws.append_row([verdict, genre, solo, team4, hit_rate, trend, notes])
   ```
6. **Git: export CSV + commit + push**

```python
# --- Git: export CSV + commit + push ---
import csv, subprocess
from pathlib import Path
from datetime import date

BASE = Path(r"C:\Organized Files\My Game Asset\Game-Research")

ratings_ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
csv_path = BASE / "data" / "genre-viability.csv"
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ratings_ws.get_all_values())

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git("add", "data/genre-viability.csv")
_git("commit", "-m", f"add: {genre} to viability ratings - {date.today().isoformat()}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
```

7. Confirm to user: verdict assigned, why, sources used.

---

## [remove] Workflow — Deleting a Genre

1. Find the row:
   ```python
   cell = ws.find(genre_name)   # finds by exact text match
   if not cell:
       # tell user genre not found
   ```
2. Confirm with the user before deleting (show the row contents).
3. Delete:
   ```python
   ws.delete_rows(cell.row)
   ```
4. **Git: export CSV + commit + push**

```python
# --- Git: export CSV + commit + push ---
import csv, subprocess
from pathlib import Path
from datetime import date

BASE = Path(r"C:\Organized Files\My Game Asset\Game-Research")

ratings_ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
csv_path = BASE / "data" / "genre-viability.csv"
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ratings_ws.get_all_values())

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git("add", "data/genre-viability.csv")
_git("commit", "-m", f"remove: {genre_name} from viability ratings - {date.today().isoformat()}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
```

5. Confirm deletion to user.

---

## Key Data Sources

- **howtomarketagame.com/blog/** — quarterly reports (Q1=Apr, Q2=Aug, Q3=Nov, Q4=Jan)
- **games-stats.com/steam/tags/** — revenue + hit rate by Steam tag
- **vginsights.com** — per-game revenue estimates
