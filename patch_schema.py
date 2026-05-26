"""
Patches competitor-lookup/SKILL.md to use the new column schema:
  Genre | Game | Year | Revenue | Review | Team size | Trailer | UI | Art style | Feature | Scope | Content | Replayability

Removed: Tier (still used internally for row coloring), Game Feel, Store URL, Date Added
Added: Genre, Trailer, UI
Renamed: Est. Revenue, Reviews, Team Size, Art Style, Features

Usage:
  python patch_schema.py
"""
import os

BASE = r"C:\Organized Files\My Game Asset\Game-Research"
TARGET = os.path.join(BASE, r".claude\skills\competitor-lookup\SKILL.md")

with open(TARGET, "r", encoding="utf-8") as f:
    content = f.read()

original = content
patches = []

# ── 1. Step 3 — Factual fields table ─────────────────────────────────────────
patches.append((
    """#### Factual fields (web search / Steam / VGInsights)

| Field | How to get it |
|---|---|
| **Game** | Name as listed on Steam |
| **Tier** | HIGH / MID / FAILURE |
| **Year** | Release year from Steam |
| **Est. Revenue** | games-stats.com, VGInsights, or Boxleiter formula: `reviews × price × 75` |
| **Reviews** | Total review count from Steam |
| **Team Size** | Solo / Duo / Team of N — check credits, itch.io, or LinkedIn. Write "Unknown" if not findable. |""",

    """#### Factual fields (web search / Steam / VGInsights)

| Field | How to get it |
|---|---|
| **Genre** | The genre being researched (e.g. "Roguelike Deckbuilder") |
| **Game** | Name as listed on Steam |
| **Year** | Release year from Steam |
| **Revenue** | games-stats.com, VGInsights, or Boxleiter formula: `reviews × price × 75` |
| **Review** | Total review count from Steam |
| **Team size** | Solo / Duo / Team of N — check credits, itch.io, or LinkedIn. Write "Unknown" if not findable. |"""
))

# ── 2. Step 3 — Descriptive fields table ──────────────────────────────────────
patches.append((
    """| Field | What to capture | Sources |
|---|---|---|
| **Art Style** | 3–6 words: rendering style, palette, visual identity. E.g. "pixel art, muted earth tones" / "hand-drawn 2D, vibrant" / "low-poly 3D minimal" | Store screenshots |
| **Game Feel** | How it feels to play: controls, feedback, pacing. Pull 1–2 player quotes. E.g. "Tight, responsive controls. 'Every hit feels satisfying' — Steam" | Steam reviews, Reddit |
| **Features** | Key mechanics and systems. List 3–5 bullet points. E.g. "Deckbuilding + perma-upgrades + branching map + boss rush; 50+ cards at launch" | Store page, description |
| **Scope** | How big the game is: estimated hours, number of levels/items/characters, EA duration if applicable. E.g. "8–15hr per run; 4 characters; 200+ items; launched in EA for 10 months" | Store page, reviews |
| **Content** | What content exists and how players perceive its depth. E.g. "3 biomes, 12 bosses, weekly challenges. 'More content than I expected for the price' — Steam" | Steam reviews, store page |
| **Replayability** | Whether players come back and why. E.g. "High — build variety strong, daily seeds, unlockable characters. 'Still playing after 80 hours' — Reddit r/roguelikes" | Steam reviews, Reddit |""",

    """| Field | What to capture | Sources |
|---|---|---|
| **Trailer** | Quality and style of the launch trailer. Does it show gameplay immediately? Note length, pacing, whether it communicates the core loop clearly. | Steam store page |
| **UI** | How clean and readable the in-game interface looks from screenshots. E.g. "Minimal HUD, clear iconography" or "Cluttered — too many overlapping panels" | Store screenshots |
| **Art style** | 3–6 words: rendering style, palette, visual identity. E.g. "pixel art, muted earth tones" / "hand-drawn 2D, vibrant" / "low-poly 3D minimal" | Store screenshots |
| **Feature** | Key mechanics and systems. List 3–5 bullet points. E.g. "Deckbuilding + perma-upgrades + branching map + boss rush; 50+ cards at launch" | Store page, description |
| **Scope** | How big the game is: estimated hours, number of levels/items/characters, EA duration if applicable. E.g. "8–15hr per run; 4 characters; 200+ items; launched in EA for 10 months" | Store page, reviews |
| **Content** | What content exists and how players perceive its depth. E.g. "3 biomes, 12 bosses, weekly challenges. 'More content than I expected for the price' — Steam" | Steam reviews, store page |
| **Replayability** | Whether players come back and why. E.g. "High — build variety strong, daily seeds, unlockable characters. 'Still playing after 80 hours' — Reddit r/roguelikes" | Steam reviews, Reddit |"""
))

# ── 3. Step 4 — Excel headers ─────────────────────────────────────────────────
patches.append((
    """headers = [
    "Game", "Tier", "Year", "Est. Revenue", "Reviews",
    "Team Size", "Art Style",
    "Game Feel", "Features", "Scope", "Content", "Replayability",
    "Store URL"
]""",

    """headers = [
    "Genre", "Game", "Year", "Revenue", "Review",
    "Team size", "Trailer", "UI", "Art style",
    "Feature", "Scope", "Content", "Replayability",
]"""
))

# ── 4. Step 4 — Excel column widths ──────────────────────────────────────────
patches.append((
    """col_widths = {
    "A": 28,   # Game
    "B": 10,   # Tier
    "C": 8,    # Year
    "D": 16,   # Est. Revenue
    "E": 10,   # Reviews
    "F": 16,   # Team Size
    "G": 22,   # Art Style
    "H": 45,   # Game Feel
    "I": 50,   # Features
    "J": 35,   # Scope
    "K": 45,   # Content
    "L": 45,   # Replayability
    "M": 18,   # Store URL
}""",

    """col_widths = {
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
}"""
))

# ── 5. Step 4 — Excel wrap text columns ──────────────────────────────────────
patches.append((
    'for col_letter in ["H", "I", "J", "K", "L"]:',
    'for col_letter in ["G", "H", "I", "J", "K", "L", "M"]:'
))

# ── 6. Step 4 — Excel example game dict ──────────────────────────────────────
patches.append((
    """    # {
    #   "game": "Slay the Spire",
    #   "tier": "HIGH",
    #   "year": 2019,
    #   "revenue": "$50M+",
    #   "reviews": "75,000+",
    #   "team": "Duo",
    #   "art_style": "hand-drawn card art, dark dungeon aesthetic",
    #   "feel": "Tight decision-making, no RNG frustration. 'Every loss feels like a lesson' — Steam",
    #   "features": "Deckbuilding + relics + map branching + 4 characters + daily climb; 500+ cards total",
    #   "scope": "5–10hr per run; 4 characters; 20+ bosses; launched EA Jan 2017, full release Jan 2019",
    #   "content": "4 characters, 20 boss types, daily challenges, custom mode. 'Endlessly replayable' — Steam",
    #   "replayability": "Very high — build variety massive, daily seeds, community mods. 'Still my go-to after 500hrs' — Reddit r/slaythespire",
    #   "store_url": "https://store.steampowered.com/app/646570/",
    # },""",

    """    # {
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
    # },"""
))

# ── 7. Step 4 — Excel row construction ───────────────────────────────────────
patches.append((
    """    row = [
        g["game"], g["tier"], g["year"], g["revenue"], g["reviews"],
        g["team"], g["art_style"],
        g["feel"], g["features"], g["scope"], g["content"], g["replayability"],
        g.get("store_url", ""),
    ]""",

    """    row = [
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ]"""
))

# ── 8. Step 4 — Excel tier color range (12 cols → 13 cols) ───────────────────
patches.append((
    '    for col in range(1, 13):  # A–L (not Store URL)',
    '    for col in range(1, 14):  # A–M'
))

# ── 9. Step 4 — Excel wrap text + top-align range ────────────────────────────
patches.append((
    '    for col in range(8, 13):  # H–L',
    '    for col in range(7, 14):  # G–M'
))

# ── 10. Step 4 — Remove Store URL hyperlink block ────────────────────────────
patches.append((
    """    # Make Store URL a hyperlink
    url = g.get("store_url", "")
    if url:
        cell = ws.cell(xl_row, 13)
        cell.hyperlink = url
        cell.value = "Steam"
        cell.font = Font(color="0563C1", underline="single")""",

    """    # (Store URL removed from schema)"""
))

# ── 11. Step 5 — Google Sheets HEADERS ───────────────────────────────────────
patches.append((
    """HEADERS = [
    "Game", "Tier", "Year", "Est. Revenue", "Reviews",
    "Team Size", "Art Style",
    "Game Feel", "Features", "Scope", "Content", "Replayability",
    "Store URL", "Date Added"
]""",

    """HEADERS = [
    "Genre", "Game", "Year", "Revenue", "Review",
    "Team size", "Trailer", "UI", "Art style",
    "Feature", "Scope", "Content", "Replayability",
]"""
))

# ── 12. Step 5 — Google Sheets column pixel widths ───────────────────────────
patches.append((
    '_col_widths = [200, 80, 60, 130, 90, 120, 160, 300, 350, 250, 300, 300, 180, 100]',
    '_col_widths = [160, 200, 60, 120, 80, 120, 200, 180, 160, 350, 250, 300, 300]'
))

# ── 13. Step 5 — Google Sheets wrap text start column (7 → 6, Trailer onward) ─
patches.append((
    '                "startColumnIndex": 7,\n                "endColumnIndex": 13,',
    '                "startColumnIndex": 6,\n                "endColumnIndex": 13,'
))

# ── 14. Step 5 — Google Sheets sheet_rows construction + tiers list ───────────
patches.append((
    """sheet_rows = []
for g in game_rows:
    sheet_rows.append([
        g["game"], g["tier"], g["year"], g["revenue"], g["reviews"],
        g["team"], g["art_style"],
        g["feel"], g["features"], g["scope"], g["content"], g["replayability"],
        g.get("store_url", ""), today,
    ])""",

    """sheet_rows = []
tiers = []
for g in game_rows:
    sheet_rows.append([
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ])
    tiers.append(g["tier"])  # tracked separately for row coloring"""
))

# ── 15. Step 5 — Google Sheets tier color lookup (row_data[1] → tiers list) ──
patches.append((
    """for i, row_data in enumerate(sheet_rows, start=start_row):
    tier = row_data[1]
    color = {"HIGH": COLOR_HIGH, "MID": COLOR_MID, "FAILURE": COLOR_FAILURE}.get(tier, COLOR_WHITE)""",

    """for i, (row_data, tier) in enumerate(zip(sheet_rows, tiers), start=start_row):
    color = {"HIGH": COLOR_HIGH, "MID": COLOR_MID, "FAILURE": COLOR_FAILURE}.get(tier, COLOR_WHITE)"""
))

# ── 16. Step 6 — Summary table column headers ────────────────────────────────
patches.append((
    '| Game | Year | Revenue | Reviews | Team | Art Style |',
    '| Genre | Game | Year | Revenue | Review | Team size | Trailer | UI | Art style |'
))

# ── Apply all patches ─────────────────────────────────────────────────────────
applied = 0
for old, new in patches:
    if old in content:
        content = content.replace(old, new)
        applied += 1
    else:
        print(f"WARNING — pattern not found:")
        print(f"  '{old[:100].strip()}...'")

if content == original:
    print("No changes made — all patterns may already be applied.")
else:
    with open(TARGET, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Done. Applied {applied}/{len(patches)} patches to competitor-lookup/SKILL.md")
    print("Re-run install-skills.ps1 and restart Claude to pick up the changes.")
