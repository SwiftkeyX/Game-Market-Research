"""One-time script: export current Google Sheets data to data/ folder for initial git commit."""
import csv
from pathlib import Path
import gspread
from google.oauth2.service_account import Credentials

BASE = Path(r"C:\Organized Files\My Game Asset\Game-Research")
CREDENTIALS_FILE = BASE / "genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1p7vPIIR8imPAZUZekbIuDW7Qfg0Jfc9xa6eEeaSv5Us/edit"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = Credentials.from_service_account_file(str(CREDENTIALS_FILE), scopes=SCOPES)
client = gspread.Client(auth=creds)
sh = client.open_by_url(SHEET_URL)

def write_csv(path, rows):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    print(f"  Exported: {path.relative_to(BASE)}")

# Export genre viability ratings
ws = sh.worksheet("Genre Viability Ratings (GO / CAUTION / AVOID)")
write_csv(BASE / "data" / "genre-viability.csv", ws.get_all_values())

# Export all competitor tabs (tabs whose title starts with 🏆)
for worksheet in sh.worksheets():
    if worksheet.title.startswith("🏆"):
        genre_display = worksheet.title.replace("🏆", "").strip()
        genre_slug = genre_display.lower().replace(" ", "-").replace("/", "-")
        write_csv(BASE / "data" / "competitors" / f"{genre_slug}.csv", worksheet.get_all_values())

print("Done.")
