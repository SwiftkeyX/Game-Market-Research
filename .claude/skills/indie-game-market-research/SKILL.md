---
name: indie-game-market-research
description: >
  Usage: (no args) = Market Overview — ranks all tracked genres by viability, hit rate, trend, revenue range.
  Usage: [genre] = Genre Deep Dive — researches 21 real games with full descriptive data.
  Examples: /indie-game-market-research  |  /indie-game-market-research [roguelike deckbuilder]
---

# Indie Game Market Research

Two modes depending on whether you provide a genre argument.

## Core Philosophy

The market is observable and rational. Study what's actually there — who's winning, who's failing,
and what separates them. Understand the landscape before you commit to anything.

---

## AUTOMATION RULES — READ FIRST

This skill runs **fully automatically** from invocation to final output. No questions, no pauses,
no confirmations. Apply these rules throughout every step:

1. **Never ask the user anything.** Not for confirmation, not for clarification, not for choices.
2. **Never pause between steps.** Run all steps sequentially without waiting for input.
3. **Ignore sub-skill blocking instructions.** If any sub-skill says "ask the user to pick",
   "confirm before continuing", or "stop and run X manually" — skip that instruction and continue.
4. **AVOID verdict = proceed anyway.** If genre-viability-check returns AVOID, note it clearly
   in the output, then run competitor-lookup regardless. The user asked for research; give them research.
5. **Ambiguous genre name = use best match.** If the genre arg doesn't exactly match a row,
   use the closest match, note it inline ("Interpreting '[input]' as '[match]'"), and proceed.
6. **Missing data = note and continue.** If a step yields incomplete results, record what was
   found, note the gap, and move on. Never stop to ask what to do.

---

## Arguments

```
/indie-game-market-research            # Mode 1: Market Overview (all genres)
/indie-game-market-research [genre]    # Mode 2: Genre Deep Dive
```

---

## Mode 1 — Market Overview (no genre provided)

**Goal**: Understand the indie game market as a whole. Which genres are healthy, risky, or saturated —
and why — so you can orient yourself before picking a direction.

### Step 1 — Load all genre ratings

Read the full ratings sheet directly using Python. Do not invoke `/genre-viability-data` interactively —
run the Python read inline:

```python
import gspread
from google.oauth2.service_account import Credentials

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1p7vPIIR8imPAZUZekbIuDW7Qfg0Jfc9xa6eEeaSv5Us/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
ws     = client.open_by_url(SHEET_URL).worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
rows   = ws.get_all_records()
for r in rows:
    print(r)
```

Use **`C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe`** to run all Python.

Note whether data is stale (compare today's date against **Next refresh due** in
`genre-viability-data/SKILL.md`). If stale, flag it in the output header — but continue regardless.

### Step 2 — Live verification pass

For any genre with stale data, or if the sheet has fewer than 10 genres, run a quick search:

```
web_search: steam indie game hit rate by genre 2025 2026 howtomarketagame
web_search: site:howtomarketagame.com games selling well 2026
```

If live data contradicts a stored verdict, update the displayed verdict and note it inline:
*"[Genre]: Updated CAUTION → GO based on live search — rising hit rate Q1 2026."*
Do not write back to the sheet; just reflect the update in the output.

### Step 3 — Output: Ranked Market Overview

Present all tracked genres grouped by verdict. Sort within each group by hit rate descending.

```
## Indie Game Market Overview
**Date**: YYYY-MM-DD  |  **Genres tracked**: N  |  **Data**: [Fresh / ⚠️ Stale — last updated DATE]

### 🟢 GO — Healthy, worth entering
| Genre | Hit Rate | Trend | Revenue Range | Key insight |
|---|---|---|---|---|
| Roguelike Deckbuilder | ~15% | Stable | $150K–$2M | Strong mid-tier, reliable demand |
| ... | | | | |

### 🟡 CAUTION — Proceed carefully
| Genre | Hit Rate | Trend | Revenue Range | Risk factor |
|---|---|---|---|---|
| Metroidvania | ~4% | Declining | $150K–$600K | Zero hits Q1 2025; Hollow Knight bar brutal |
| ... | | | | |

### 🔴 AVOID — Saturated or declining
| Genre | Hit Rate | Trend | Revenue Range | Why to avoid now |
|---|---|---|---|---|
| Bullet Heaven Clone | <1% | Saturated | $50K–$300K | Only 1 hit all of 2024 |
| ... | | | | |

---
**Key pattern**: [1–2 sentences on the most important signal in the current market snapshot]
**Data freshness**: [N genres verified live today | data from sheet cache dated DATE]
```

End with a tip (no question, no prompt):
> *To dive deeper into any genre: `/indie-game-market-research [genre name]`*

---

## Mode 2 — Genre Deep Dive (genre provided)

**Goal**: Understand one specific genre in depth — who's in it, what they built, what the market
looks like at each tier, and what patterns separate winners from failures.

### Step 1 — Resolve genre name

Parse the genre from the argument. If it doesn't exactly match a row in the sheet, find the closest
match by name and proceed with it. Note the interpretation inline if different from the input:
*"Interpreting 'deckbuilder' as 'Roguelike Deckbuilder'."*

Do not ask the user to confirm the match. Just proceed.

### Step 2 — Quick viability snapshot

Read the viability data directly from the sheet (same Python read as Mode 1, Step 1).
Find the genre row and extract: Verdict, Hit Rate, Trend, Solo Target, Team of 4 Target, Notes.

Run a live verification search:
```
web_search: site:howtomarketagame.com [genre] 2026
web_search: [genre] steam indie revenue hit rate 2025 2026
```

Display as a compact header:
```
**[Genre]** — Verdict: GO / CAUTION / AVOID  |  Hit Rate: X%  |  Trend: [trend]
[One sentence from live search: Confirmed / Updated — what the web shows]
```

If verdict is **AVOID**: display the warning, then **immediately continue to Step 3 anyway**.
The user explicitly invoked a deep dive — always deliver it.

### Step 3 — Research 21 games

Invoke `/competitor-lookup [genre]` as a sub-step. It handles everything:
finding candidates, categorizing tiers, collecting all data fields, writing the Excel file,
and writing to the Google Sheets tab. Incorporate its output into Step 4 below.

### Step 4 — Market Shape Summary

After all 21 games are collected, produce this summary:

```
## Genre Deep Dive: [Genre]
**Verdict**: GO / CAUTION / AVOID  |  **Hit Rate**: X%  |  **Trend**: [trend]
**Date**: YYYY-MM-DD  |  **Games researched**: 21 (7 HIGH / 7 MID / 7 FAILURE)

### What's winning in this space
[2–3 sentences: what do HIGH-tier games share? Look at art style, scope, features, team size.]

### What's failing
[2–3 sentences: common patterns in FAILURE games — what's missing or wrong?]

### Market shape at a glance
| | HIGH | MID | FAILURE |
|---|---|---|---|
| Revenue range | $X–$Y | $A–$B | <$C |
| Typical team size | [pattern] | [pattern] | [pattern] |
| Dominant art styles | [styles] | [styles] | [styles] |
| Common features | [features] | [features] | [features] |
| Typical scope | [scope] | [scope] | [scope] |

### Replayability & content pattern
[1–2 sentences: what do HIGH games offer in terms of replay value and content depth?]

---
**Full data**: gameplay-review-[genre-slug].xlsx
**Sheet tab**: 🏆 [genre-slug] (Google Sheets)
```

---

## Key Data Sources

- **games-stats.com/steam/?tag=<tag>** — Revenue estimates by Steam tag
- **howtomarketagame.com** — Annual genre hit-rate analysis (Chris Zukowski)
- **vginsights.com** — Revenue estimates for specific games
- **store.steampowered.com** — Store pages, review counts, pricing

---

## Reference Files

- `references/subgenre-guide.md` — Per-subgenre breakdown with examples
- `references/competitor-examples.md` — Pre-researched data for common genres
- `references/quality-benchmark-guide.md` — Evidence sources for each data dimension
