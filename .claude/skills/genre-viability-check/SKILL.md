---
name: genre-viability-check
description: >
  Usage: [genre] — Look up a genre's current GO/CAUTION/AVOID rating, hit rate, and trend.
  Verifies with a live web search. Returns a verdict with confidence level.
  Example: [roguelike deckbuilder]
---

# Genre Viability Check

Checks whether a genre is currently a good bet on Steam. Reads the live genre-viability-data
sheet, verifies with a web search, and returns a structured verdict.

## Arguments

```
/genre-viability-check [genre]   # check a specific genre
/genre-viability-check           # list all GO genres and ask user to pick
```

---

## Workflow

### Step 1 — Read genre-viability-data sheet

Run Python to fetch the current rating from the Google Sheet:

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
```

**If genre was provided (arg 1):** find the row whose `Genre` column is a case-insensitive match.
If no exact match, look for partial matches and show up to 3 closest options.

**If no genre was provided:** filter `rows` to `Verdict == "GO"`, sort by hit rate descending,
display the top 5 with their hit rates and trends. Ask the user to pick one, then continue with
that genre.

**Staleness check:** the skill file has a `Last updated` date at the top of `genre-viability-data`.
If today's date is past `Next refresh due`, warn the user:
> ⚠️ Genre data is overdue for refresh. Run `/genre-viability-data [update]` to get current numbers.
> Showing last known data — verify with live search below.

---

### Step 2 — Live verification

Always run these two web searches, regardless of whether the genre was found in the sheet:

```
web_search: site:howtomarketagame.com [genre] 2026
web_search: [genre] steam indie revenue hit rate 2025 2026
```

Read the results. Compare against the sheet rating:

- **Confirms rating**: no change needed — note the source
- **Contradicts rating** (e.g. sheet says GO but new data shows declining hits): flag the discrepancy,
  use the live data as the operative verdict, and suggest running `/genre-viability-data [update]`
- **No live data found**: proceed with sheet data, note that verification was inconclusive

---

### Step 3 — Output

```
## Genre: [Genre Name]

**Verdict**: GO / CAUTION / AVOID
**Hit Rate**: X% of releases reach 1000+ reviews (~$150K revenue)
**Trend**: Growing / Stable / Declining / Emerging

**Solo target**: $X – $Y
**Team of 4 target**: $X – $Y

**Notes**: [key considerations from the sheet notes field]

**Live verification**: [Confirmed / Updated / Inconclusive] — [1-sentence summary of what the web search found]
**Data as of**: [date from sheet]
```

If verdict is **AVOID**, add a bolded warning:
> ⚠️ This genre has near-zero hit rate. Proceeding with market research is not recommended.
> If you want to continue anyway, run `/competitor-lookup [genre]`.

If verdict is **CAUTION**, add:
> ⚠️ Proceed with caution. This genre is viable but has meaningful risk factors noted above.
