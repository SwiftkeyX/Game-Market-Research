---
name: competitor-lookup
description: "INTERNAL — Do not invoke directly."
---

## ⚠️ INTERNAL SKILL

This skill is invoked automatically by `/indie-game-market-research [genre]`.
If the **user** typed `/competitor-lookup` directly, respond with this message and stop:

> "competitor-lookup is an internal sub-skill. To research a genre, run:
> `/indie-game-market-research [genre name]`
> It runs competitor-lookup automatically as part of the full workflow."

Do not execute any further steps in this skill when invoked directly by the user.

---

# Competitor Lookup

Researches the real game landscape for a genre. Collects both factual data (revenue, reviews,
team size, year) and descriptive intelligence (art style, game feel, key features, scope,
content depth, replayability) for 21 games across three tiers.

**Data philosophy**: Describe what's actually there. No scores, no pass/fail judgments.
The goal is to understand the market as it is — what games exist, what they offer, and what
patterns emerge at each tier.

## Arguments

```
/competitor-lookup [genre]   # required — the genre to research
```

---

## Workflow

### Step 1 — Find candidates

Run web searches to build a pool of released games in the genre:

```
web_search: [genre] steam games released 2022 2023 2024 2025
web_search: [genre] indie steam most reviewed
web_search: best [genre] games steam 2024 2025
web_search: [genre] steam games low reviews failed
web_search: site:howtomarketagame.com [genre] examples
```

Also check games-stats.com for the tag if possible:
```
web_fetch: https://games-stats.com/steam/?tag=[genre-tag]
```
If 403, use:
```
web_search: games-stats.com [genre] median revenue reviews
```

Aim for 25–30 candidates so you can select the best 7 per tier.

### Step 2 — Categorize into tiers

For each candidate, look up review count on Steam or via web search:

| Tier | Criteria |
|---|---|
| 🟢 HIGH | 1000+ reviews — clearly profitable, positive reception |
| 🟡 MID | 100–999 reviews — alive but modest |
| 🔴 FAILURE | <100 reviews — abandoned, mixed/negative, or dead playerbase |

Select 7 per tier (21 total). If fewer than 7 exist in a tier, use all you can find.

**Prioritize recency**: games from the last 3–4 years. Include older titles only if they
define the genre or are instructive cautionary tales.

**Include at least one solo dev** in HIGH or MID — important market signal.

---

### Step 3 — Collect data for each game

For each of the 21 games, collect the following fields. Use the sources listed for each.

#### Factual fields (web search / Steam / VGInsights)

| Field | How to get it |
|---|---|
| **Genre** | The genre being researched (e.g. "Roguelike Deckbuilder") |
| **Game** | Name as listed on Steam |
| **Year** | Release year from Steam |
| **Revenue** | games-stats.com, VGInsights, or Boxleiter formula: `reviews × price × 75` |
| **Review** | Total review count from Steam |
| **Team size** | Solo / Duo / Team of N — check credits, itch.io, or LinkedIn. Write "Unknown" if not findable. |

#### Descriptive fields (store page + player reviews)

| Field | What to capture | Sources |
|---|---|---|
| **Trailer** | Quality and style of the launch trailer. Does it show gameplay immediately? Note length, pacing, whether it communicates the core loop clearly. | Steam store page |
| **UI** | How clean and readable the in-game interface looks from screenshots. E.g. "Minimal HUD, clear iconography" or "Cluttered — too many overlapping panels" | Store screenshots |
| **Art style** | 3–6 words: rendering style, palette, visual identity. E.g. "pixel art, muted earth tones" / "hand-drawn 2D, vibrant" / "low-poly 3D minimal" | Store screenshots |
| **Feature** | Key mechanics and systems. List 3–5 bullet points. E.g. "Deckbuilding + perma-upgrades + branching map + boss rush; 50+ cards at launch" | Store page, description |
| **Scope** | How big the game is: estimated hours, number of levels/items/characters, EA duration if applicable. E.g. "8–15hr per run; 4 characters; 200+ items; launched in EA for 10 months" | Store page, reviews |
| **Content** | What content exists and how players perceive its depth. E.g. "3 biomes, 12 bosses, weekly challenges. 'More content than I expected for the price' — Steam" | Steam reviews, store page |
| **Replayability** | Whether players come back and why. E.g. "High — build variety strong, daily seeds, unlockable characters. 'Still playing after 80 hours' — Reddit r/roguelikes" | Steam reviews, Reddit |

**Search pattern per game:**
```
web_search: "[game name]" steam review gameplay feel controls
web_search: "[game name]" steam review hours content replayability
web_search: site:reddit.com "[game name]" review impressions
web_search: "[game name]" vginsights revenue estimate
```

Batch these per game to be efficient. Prioritize HIGH-tier games — if time is short,
collect lighter data for FAILURE games.

---

### Step 4 — Save Excel file

Install openpyxl if needed:
```
C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe -m pip install openpyxl
```

Save to: `C:\Organized Files\My Game Asset\Game Market\Game-Research\exports\gameplay-review-[genre-slug].xlsx`

```python
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="CFE2F3")
HIGH_FILL   = PatternFill("solid", fgColor="D9EAD3")
MID_FILL    = PatternFill("solid", fgColor="FFF2CC")
FAIL_FILL   = PatternFill("solid", fgColor="FFE6E6")

TIER_FILLS  = {"HIGH": HIGH_FILL, "MID": MID_FILL, "FAILURE": FAIL_FILL}

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Genre Deep Dive"

headers = [
    "Genre", "Game", "Year", "Revenue", "Review",
    "Team size", "Trailer", "UI", "Art style",
    "Feature", "Scope", "Content", "Replayability",
]
ws.append(headers)
for cell in ws[1]:
    cell.fill = HEADER_FILL
    cell.font = Font(bold=True)

col_widths = {
    "A": 20,   # Genre
    "B": 28,   # Game
    "C": 8,    # Year
    "D": 16,   # Revenue
    "E": 10,   # Review
    "F": 16,   # Team size
    "G": 30,   # Trailer
    "H": 25,   # UI
    "I": 22,   # Art style
    "J": 50,   # Feature
    "K": 35,   # Scope
    "L": 45,   # Content
    "M": 45,   # Replayability
}
for col_letter, width in col_widths.items():
    ws.column_dimensions[col_letter].width = width

# Wrap text for descriptive columns
for col_letter in ["G", "H", "I", "J", "K", "L", "M"]:
    ws.column_dimensions[col_letter].width = col_widths[col_letter]

# --- Replace with actual game data ---
game_rows = [
    # dict per game, e.g.:
    # {
    #   "genre": "Roguelike Deckbuilder",
    #   "game": "Slay the Spire",
    #   "tier": "HIGH",   # internal — used for row color only, not a column
    #   "year": 2019,
    #   "revenue": "$50M+",
    #   "review": "75,000+",
    #   "team_size": "Duo",
    #   "trailer": "60s, shows combat loop in first 5s, strong music and pacing",
    #   "ui": "Clean card layout, minimal clutter, readable stat numbers",
    #   "art_style": "hand-drawn card art, dark dungeon aesthetic",
    #   "feature": "Deckbuilding + relics + map branching + 4 characters + daily climb; 500+ cards total",
    #   "scope": "5–10hr per run; 4 characters; 20+ bosses; launched EA Jan 2017, full release Jan 2019",
    #   "content": "4 characters, 20 boss types, daily challenges, custom mode. 'Endlessly replayable' — Steam",
    #   "replayability": "Very high — build variety massive, daily seeds, community mods. 'Still my go-to after 500hrs' — Reddit r/slaythespire",
    # },
]

for g in game_rows:
    row = [
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ]
    ws.append(row)
    xl_row = ws.max_row

    # Color row by tier
    fill = TIER_FILLS.get(g["tier"])
    if fill:
        for col in range(1, 14):  # A–M
            ws.cell(xl_row, col).fill = fill

    # Wrap text + top-align descriptive columns
    for col in range(7, 14):  # G–M
        cell = ws.cell(xl_row, col)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # (Store URL removed from schema)

# Freeze header row
ws.freeze_panes = "A2"

genre_slug = genre.lower().replace(" ", "-").replace("/", "-")
output_path = rf"C:\Organized Files\My Game Asset\Game Market\Game-Research\exports\gameplay-review-{genre_slug}.xlsx"
wb.save(output_path)
print(f"Saved: {output_path}")

# --- Snapshot archive ---
import shutil
from pathlib import Path

snapshot_dir = Path(rf"C:\Organized Files\My Game Asset\Game Market\Game-Research\snapshots") / genre_slug
snapshot_dir.mkdir(parents=True, exist_ok=True)
snapshot_path = snapshot_dir / f"{date.today().isoformat()}.xlsx"
shutil.copy2(output_path, snapshot_path)
print(f"Snapshot: {snapshot_path}")

# --- Update research-log.md ---
log_path = Path(r"C:\Organized Files\My Game Asset\Game Market\Game-Research\research-log.md")
today_str = date.today().isoformat()
genre_display = " ".join(w.capitalize() for w in genre.split())
high_count = sum(1 for g in game_rows if g["tier"] == "HIGH")
mid_count  = sum(1 for g in game_rows if g["tier"] == "MID")
fail_count = sum(1 for g in game_rows if g["tier"] == "FAILURE")
entry = f"| {today_str} | {genre_display} | {high_count} | {mid_count} | {fail_count} | `snapshots/{genre_slug}/{today_str}.xlsx` |\n"
if log_path.exists():
    log_path.write_text(log_path.read_text() + entry, encoding="utf-8")
else:
    log_path.write_text(
        "# Research Log\n\n"
        "| Date | Genre | HIGH | MID | FAILURE | Snapshot |\n"
        "|------|-------|------|-----|---------|----------|\n"
        + entry,
        encoding="utf-8",
    )
print(f"research-log.md updated.")
```

---

### Step 5 — Write to Google Sheets

Write all 21 games to a **genre-specific tab** in the Google Sheet (e.g. `🏆 Roguelike Deckbuilder`).
Using a dedicated tab per genre keeps each research session clean and avoids schema conflicts.
Each run **replaces** the previous tab entirely — no archive tabs are kept. Historical comparison is done via `git diff` on the exported CSV.

```python
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import time

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game Market\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = [
    "Genre", "Game", "Year", "Revenue", "Review",
    "Team size", "Trailer", "UI", "Art style",
    "Feature", "Scope", "Content", "Replayability",
]

COLOR_HEADER  = {"red": 0.812, "green": 0.886, "blue": 0.953}
COLOR_HIGH    = {"red": 0.851, "green": 0.918, "blue": 0.827}
COLOR_MID     = {"red": 1.0,   "green": 0.949, "blue": 0.800}
COLOR_FAILURE = {"red": 1.0,   "green": 0.902, "blue": 0.902}
COLOR_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

today = date.today().isoformat()
genre_slug = genre.lower().replace(" ", "-").replace("/", "-")
# genre_display must use title-cased words, never a slug (e.g. "Roguelike Deckbuilder" not "roguelike-deckbuilder")
genre_display = " ".join(w.capitalize() for w in genre.split())
tab_name = f"🏆 {genre_display}"

# No archive tabs — historical comparison is done via git diff on the exported CSV.
# Delete the existing active tab for this genre (new-format: 🏆 prefix)
try:
    old_ws = sh.worksheet(tab_name)
    sh.del_worksheet(old_ws)
except gspread.WorksheetNotFound:
    pass

# Delete 🏆-prefixed tab using old slug name e.g. "🏆 roguelike-deckbuilder"
slug_tab_name = f"🏆 {genre_slug}"
if slug_tab_name != tab_name:
    try:
        sh.del_worksheet(sh.worksheet(slug_tab_name))
    except gspread.WorksheetNotFound:
        pass

# Delete 🏆-prefixed tab with no space after emoji e.g. "🏆Roguelike" (old upload script convention)
nospace_tab_name = f"🏆{genre_display}"
if nospace_tab_name != tab_name:
    try:
        sh.del_worksheet(sh.worksheet(nospace_tab_name))
    except gspread.WorksheetNotFound:
        pass

# Delete old-format tabs for this genre (no emoji prefix, dated or plain, from older runs)
import re
old_format_pattern = re.compile(
    r"^" + re.escape(genre_display) + r"(\s+\(\d{4}-\d{2}-\d{2}\))?$",
    re.IGNORECASE
)
for t in sh.worksheets():
    if old_format_pattern.match(t.title):
        sh.del_worksheet(t)

# Delete broken-emoji orphan tabs for this genre (case-insensitive, e.g. "?? Narrative RPG")
genre_lower = genre_display.lower()
for t in sh.worksheets():
    if genre_lower in t.title.lower() and not t.title.startswith("🏆"):
        sh.del_worksheet(t)

ws = sh.add_worksheet(title=tab_name, rows=500, cols=len(HEADERS) + 2)
ws.append_row(HEADERS)
ws.format(f"A1:{chr(64 + len(HEADERS))}1", {
    "backgroundColor": COLOR_HEADER,
    "textFormat": {"bold": True}
})

# Set column widths, wrap text for descriptive columns, freeze header row
_col_widths = [160, 200, 60, 120, 80, 120, 200, 180, 160, 350, 250, 300, 300]
sh.batch_update({"requests": [
    {
        "updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }
    for i, px in enumerate(_col_widths)
] + [
    {
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 1,
                "startColumnIndex": 6,
                "endColumnIndex": 13,
            },
            "cell": {"userEnteredFormat": {
                "wrapStrategy": "WRAP",
                "verticalAlignment": "TOP",
            }},
            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",
        }
    },
    {
        "updateSheetProperties": {
            "properties": {
                "sheetId": ws.id,
                "gridProperties": {"frozenRowCount": 1},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    },
]})

# Build rows
sheet_rows = []
tiers = []
for g in game_rows:
    sheet_rows.append([
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ])
    tiers.append(g["tier"])  # tracked separately for row coloring

ws.append_rows(sheet_rows, value_input_option="USER_ENTERED")

# Auto-resize rows to show all wrapped text
sh.batch_update({"requests": [{
    "autoResizeDimensions": {
        "dimensions": {
            "sheetId": ws.id,
            "dimension": "ROWS",
            "startIndex": 1,
        }
    }
}]})

# Color rows by tier
all_vals = ws.get_all_values()
total_rows = len(all_vals)
start_row = total_rows - len(sheet_rows) + 1

format_requests = []
for i, (row_data, tier) in enumerate(zip(sheet_rows, tiers), start=start_row):
    color = {"HIGH": COLOR_HIGH, "MID": COLOR_MID, "FAILURE": COLOR_FAILURE}.get(tier, COLOR_WHITE)
    format_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": i - 1,
                "endRowIndex": i,
                "startColumnIndex": 0,
                "endColumnIndex": len(HEADERS),
            },
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    })
    if len(format_requests) >= 25:
        sh.batch_update({"requests": format_requests})
        format_requests = []
        time.sleep(1.2)

if format_requests:
    sh.batch_update({"requests": format_requests})

print(f"Written {len(sheet_rows)} games to tab '{tab_name}'.")
```

---

### Step 6.5 — Git: export CSV + commit + push

Export the tab just written to a diffable CSV, then commit and push to GitHub.

```python
# --- Git: export CSV + commit + push ---
import csv, subprocess
from pathlib import Path

BASE = Path(r"C:\Organized Files\My Game Asset\Game Market\Game-Research")

data_dir = BASE / "data" / "competitors"
data_dir.mkdir(parents=True, exist_ok=True)
csv_path = data_dir / f"{genre_slug}.csv"
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ws.get_all_values())
print(f"Exported: {csv_path.relative_to(BASE)}")

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

today_str = date.today().isoformat()
_git("add", f"data/competitors/{genre_slug}.csv", "research-log.md")
_git("commit", "-m", f"research: {genre_display} competitor data - {today_str}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
```

`ws`, `genre_slug`, `genre_display`, and `date` are all already in scope from Step 5.

---

### Step 6 — Text summary output

Display a grouped summary in chat after writing:

```
## Competitor Research: [Genre]
**Games found**: 7 HIGH / 7 MID / 7 FAILURE  |  **Date**: YYYY-MM-DD

### 🟢 HIGH (1000+ reviews)
| Genre | Game | Year | Revenue | Review | Team size | Trailer | UI | Art style |
|---|---|---|---|---|---|
| ... | | | | | |

### 🟡 MID (100–999 reviews)
| Genre | Game | Year | Revenue | Review | Team size | Trailer | UI | Art style |
...

### 🔴 FAILURE (<100 reviews)
| Genre | Game | Year | Revenue | Review | Team size | Trailer | UI | Art style |
...

**Solo dev benchmark**: [Game] — [revenue], [reviews] reviews
**Excel file**: gameplay-review-[genre].xlsx
**Snapshot**: snapshots/[genre-slug]/YYYY-MM-DD.xlsx
**Sheet tab**: 🏆 [Genre Name] (title case)  _(previous tab replaced; use `git diff HEAD~1 data/competitors/[genre-slug].csv` to compare runs)_
```
