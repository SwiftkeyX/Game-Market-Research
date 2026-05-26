"""
Lists all tabs in the Google Sheet, then lets you choose which to delete.
Usage: python clean_tabs.py
"""
import gspread
from google.oauth2.service_account import Credentials

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

tabs = sh.worksheets()

print("Current tabs in your Google Sheet:")
print()
for i, ws in enumerate(tabs):
    print(f"  [{i}] {ws.title}")

print()
print("Enter the numbers of tabs to DELETE (comma-separated), or press Enter to cancel:")
raw = input("> ").strip()

if not raw:
    print("Cancelled — nothing deleted.")
else:
    indices = [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]
    to_delete = [tabs[i] for i in indices if 0 <= i < len(tabs)]

    print()
    print("About to delete:")
    for ws in to_delete:
        print(f"  - {ws.title}")
    confirm = input("Type YES to confirm: ").strip()

    if confirm == "YES":
        for ws in to_delete:
            sh.del_worksheet(ws)
            print(f"  Deleted: {ws.title}")
        print("Done.")
    else:
        print("Cancelled.")
