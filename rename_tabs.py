"""
Run this once to fix all Google Sheets skill issues:
  1. Rename the 3 core Google Sheet tabs to clearer names
  2. Update all skill files so they reference the new tab names
  3. Fix genre tab naming (title case instead of lowercase-dashes)
  4. Add wrap text + row auto-resize to descriptive columns
  5. Add header row freeze
  6. Retire the deprecated gameplay-review skill file
  7. Simplify indie-game-market-research Step 3 (no more duplicated instructions)

Usage:
  python rename_tabs.py
"""
import os
import gspread
from google.oauth2.service_account import Credentials

# ── Config ────────────────────────────────────────────────────────────────────

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1p7vPIIR8imPAZUZekbIuDW7Qfg0Jfc9xa6eEeaSv5Us/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
BASE   = r"C:\Organized Files\My Game Asset\Game-Research"

# ── Step 1: Rename core Google Sheets tabs ────────────────────────────────────

RENAMES = {
    "Current Ratings":          "Genre Viability Ratings (GO / CAUTION / AVOID)",
    "Current Rating":           "Genre Viability Ratings (GO / CAUTION / AVOID)",
    "Index":                    "Research Session History",
    "\U0001f4cb Index":         "Research Session History",
    "Competitors":              "All Competitor Games by Genre",
    "\U0001f3c6 Competitors":   "All Competitor Games by Genre",
}

print("=" * 60)
print("STEP 1 — Renaming core Google Sheets tabs")
print("=" * 60)

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

tabs = sh.worksheets()
print(f"Found {len(tabs)} tab(s): {[t.title for t in tabs]}\n")

renamed = 0
for ws in tabs:
    old = ws.title
    if old in RENAMES:
        new = RENAMES[old]
        ws.update_title(new)
        print(f"  Renamed: '{old}'  ->  '{new}'")
        renamed += 1

print(f"\n  Done. {renamed} tab(s) renamed." if renamed else "  No tabs matched — may already be renamed.")

# ── Steps 2–7: Patch skill files ─────────────────────────────────────────────

print()
print("=" * 60)
print("STEPS 2-7 — Patching skill files")
print("=" * 60)

# Each entry: (relative path, [(old_string, new_string), ...])
FILE_PATCHES = [

    # ── genre-viability-data: tab name references ─────────────────────────────
    (r".claude\skills\genre-viability-data\SKILL.md", [
        (
            'ws = sh.worksheet("Current Ratings")\n    ws.clear()\nexcept gspread.WorksheetNotFound:\n    ws = sh.add_worksheet("Current Ratings", rows=100, cols=10)',
            'ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")\n    ws.clear()\nexcept gspread.WorksheetNotFound:\n    ws = sh.add_worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)", rows=100, cols=10)',
        ),
        (
            "print(f\"Done. Written {len(GENRES)} genres to 'Current Ratings' and 'v1 - May 2026'.\")",
            "print(f\"Done. Written {len(GENRES)} genres to 'Genre Viability Ratings (GO / CAUTION / AVOID)' and 'v1 - May 2026'.\")",
        ),
        (
            'ws     = client.open_by_url(SHEET_URL).worksheet("Current Ratings")',
            'ws     = client.open_by_url(SHEET_URL).worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")',
        ),
        (
            'ws_id   = sh.worksheet("Current Ratings").id',
            'ws_id   = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)").id',
        ),
    ]),

    # ── genre-viability-check: tab name reference ─────────────────────────────
    (r".claude\skills\genre-viability-check\SKILL.md", [
        (
            'ws     = client.open_by_url(SHEET_URL).worksheet("Current Ratings")',
            'ws     = client.open_by_url(SHEET_URL).worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")',
        ),
    ]),

    # ── indie-game-market-research: tab name + simplify Step 3 ───────────────
    (r".claude\skills\indie-game-market-research\SKILL.md", [

        # Tab name reference
        (
            'ws     = client.open_by_url(SHEET_URL).worksheet("Current Ratings")',
            'ws     = client.open_by_url(SHEET_URL).worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")',
        ),

        # Simplify Step 3 — replace duplicated inline instructions with a clean sub-skill call
        (
            "### Step 3 — Research 21 games\n\nRun the full competitor research inline (do not sub-invoke `/competitor-lookup` interactively).\nFollow the exact workflow from `competitor-lookup/SKILL.md`:\n\n1. Find 25–30 candidate games via web search\n2. Categorize into HIGH (1000+ reviews) / MID (100–999) / FAILURE (<100)\n3. Select 7 per tier (21 total)\n4. For each game collect:\n   - **Factual**: Game, Tier, Year, Est. Revenue, Reviews, Team Size, Art Style\n   - **Descriptive** (from store page + player reviews):\n     Game Feel, Features, Scope, Content, Replayability\n5. Write to Excel file: `gameplay-review-[genre-slug].xlsx`\n6. Write to Google Sheets tab: `\U0001f3c6 [genre-slug]`\n\nFull data collection and output instructions are in `competitor-lookup/SKILL.md`.",
            "### Step 3 — Research 21 games\n\nInvoke `/competitor-lookup [genre]` as a sub-step. It handles everything:\nfinding candidates, categorizing tiers, collecting all data fields, writing the Excel file,\nand writing to the Google Sheets tab. Incorporate its output into Step 4 below.",
        ),
    ]),

    # ── competitor-lookup: tab naming + wrap text + freeze + auto-resize ──────
    (r".claude\skills\competitor-lookup\SKILL.md", [

        # Fix 1 — genre tab name: lowercase-dashes -> Title Case with spaces
        (
            'genre_slug = genre.lower().replace(" ", "-").replace("/", "-")\ntab_name = f"\U0001f3c6 {genre_slug}"',
            'genre_slug = genre.lower().replace(" ", "-").replace("/", "-")\ngenre_display = " ".join(w.capitalize() for w in genre.split())\ntab_name = f"\U0001f3c6 {genre_display}"',
        ),

        # Fix 2 — extend the column-width batch_update to also add wrap text + freeze
        (
            '# Set column widths\n_col_widths = [200, 80, 60, 130, 90, 120, 160, 300, 350, 250, 300, 300, 180, 100]\nsh.batch_update({"requests": [\n    {\n        "updateDimensionProperties": {\n            "range": {"sheetId": ws.id, "dimension": "COLUMNS",\n                      "startIndex": i, "endIndex": i + 1},\n            "properties": {"pixelSize": px},\n            "fields": "pixelSize",\n        }\n    }\n    for i, px in enumerate(_col_widths)\n]})',
            '# Set column widths, wrap text for descriptive columns, freeze header row\n_col_widths = [200, 80, 60, 130, 90, 120, 160, 300, 350, 250, 300, 300, 180, 100]\nsh.batch_update({"requests": [\n    {\n        "updateDimensionProperties": {\n            "range": {"sheetId": ws.id, "dimension": "COLUMNS",\n                      "startIndex": i, "endIndex": i + 1},\n            "properties": {"pixelSize": px},\n            "fields": "pixelSize",\n        }\n    }\n    for i, px in enumerate(_col_widths)\n] + [\n    {\n        "repeatCell": {\n            "range": {\n                "sheetId": ws.id,\n                "startRowIndex": 1,\n                "startColumnIndex": 7,\n                "endColumnIndex": 13,\n            },\n            "cell": {"userEnteredFormat": {\n                "wrapStrategy": "WRAP",\n                "verticalAlignment": "TOP",\n            }},\n            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",\n        }\n    },\n    {\n        "updateSheetProperties": {\n            "properties": {\n                "sheetId": ws.id,\n                "gridProperties": {"frozenRowCount": 1},\n            },\n            "fields": "gridProperties.frozenRowCount",\n        }\n    },\n]})',
        ),

        # Fix 3 — add row auto-resize after appending rows
        (
            'ws.append_rows(sheet_rows, value_input_option="USER_ENTERED")',
            'ws.append_rows(sheet_rows, value_input_option="USER_ENTERED")\n\n# Auto-resize rows to show all wrapped text\nsh.batch_update({"requests": [{\n    "autoResizeDimensions": {\n        "dimensions": {\n            "sheetId": ws.id,\n            "dimension": "ROWS",\n            "startIndex": 1,\n        }\n    }\n}]})',
        ),

        # Fix 4 — update the summary footer to show Title Case tab name
        (
            '**Sheet tab**: \U0001f3c6 [genre-slug]',
            '**Sheet tab**: \U0001f3c6 [Genre Name] (title case)',
        ),
    ]),

    # ── sheet-column-guide: update tab headings in documentation ──────────────
    (r"references\sheet-column-guide.md", [
        (
            "## Tab: Current Ratings (and v1 / v2 snapshots)",
            "## Tab: Genre Viability Ratings (GO / CAUTION / AVOID) (and v1 / v2 snapshots)",
        ),
        (
            "## Tab: \U0001f4cb Index",
            "## Tab: Research Session History",
        ),
        (
            "## Tab: \U0001f3c6 Competitors",
            "## Tab: All Competitor Games by Genre",
        ),
        (
            "## Tab: \U0001f4ca Quality Scores",
            "## Tab: Quality Scores",
        ),
    ]),
]

for rel_path, patches in FILE_PATCHES:
    full_path = os.path.join(BASE, rel_path)
    if not os.path.exists(full_path):
        print(f"  SKIP (not found): {rel_path}")
        continue

    with open(full_path, "r", encoding="utf-8") as f:
        content = f.read()

    updated = content
    applied = 0
    for old, new in patches:
        if old in updated:
            updated = updated.replace(old, new)
            applied += 1
        else:
            print(f"  WARNING — pattern not found in {rel_path}:")
            print(f"    '{old[:80]}...'")

    if updated == content:
        print(f"  — No changes needed: {rel_path}")
    else:
        with open(full_path, "w", encoding="utf-8") as f:
            f.write(updated)
        print(f"  Updated ({applied} fix(es)): {rel_path}")

# ── Step 6: Retire deprecated gameplay-review skill ──────────────────────────

print()
print("=" * 60)
print("STEP 6 — Retiring deprecated gameplay-review skill")
print("=" * 60)

deprecated_path = os.path.join(BASE, r".claude\skills\gameplay-review\SKILL.md")
retired_path    = deprecated_path + ".retired"

if os.path.exists(retired_path):
    print("  Already retired — nothing to do.")
elif os.path.exists(deprecated_path):
    os.rename(deprecated_path, retired_path)
    print(f"  Renamed SKILL.md -> SKILL.md.retired")
    print(f"  The skill can no longer be invoked, but the file is preserved.")
else:
    print(f"  SKIP — file not found: {deprecated_path}")

print()
print("All done. You can delete this script.")
