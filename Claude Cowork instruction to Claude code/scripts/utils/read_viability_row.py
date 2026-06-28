"""
read_viability_row.py
=====================
Reads the Genre Viability Ratings sheet and prints the row for
Roguelike Deckbuilder (or any genre you specify below).

Usage:
  C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe read_viability_row.py

Requirements (install once if needed):
  C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe -m pip install gspread google-auth
"""

import gspread
from google.oauth2.service_account import Credentials

# ── Config ─────────────────────────────────────────────────────────────────────
CREDENTIALS_FILE = r"C:\Organized Files\Working\Agent\Game-Market-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL        = "https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit"
SCOPES           = ["https://www.googleapis.com/auth/spreadsheets"]
WORKSHEET_NAME   = "Genre Viability Ratings (GO / CAUTION / AVOID)"

SEARCH_GENRE     = "roguelike deckbuilder"   # ← change this to look up a different genre

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    print("Connecting to Google Sheets...")
    creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.Client(auth=creds)
    ws     = client.open_by_url(SHEET_URL).worksheet(WORKSHEET_NAME)

    rows = ws.get_all_records()
    print(f"Loaded {len(rows)} rows from '{WORKSHEET_NAME}'.\n")

    matches = [r for r in rows if SEARCH_GENRE.lower() in str(r).lower()]

    if not matches:
        print(f"No row found matching '{SEARCH_GENRE}'.")
        print("\nAll genres in sheet:")
        for r in rows:
            # Print first column value (usually the genre name)
            first_val = list(r.values())[0] if r else "(empty)"
            print(f"  {first_val}")
    else:
        for r in matches:
            print("─" * 60)
            for key, val in r.items():
                print(f"  {key}: {val}")
        print("─" * 60)


if __name__ == "__main__":
    main()
