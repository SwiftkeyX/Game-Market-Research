import gspread
from google.oauth2.service_account import Credentials
import csv, subprocess, sys
from pathlib import Path
from datetime import date

sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = r'C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
BASE = Path(r'C:\Organized Files\My Game Asset\Game-Research')

creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh = client.open_by_url(SHEET_URL)
ws = sh.worksheet('Genre Viability Ratings (GO / CAUTION / AVOID)')

# Check for duplicate
rows = ws.get_all_records()
if any(r['Genre'].lower() == 'roguelike' for r in rows):
    print('Roguelike is already tracked. Use /genre-viability-data to update it.')
    sys.exit(0)

# New row
verdict     = 'CAUTION'
genre       = 'Roguelike'
solo        = '$80K-$350K'
team4       = '$300K-$1.5M'
hit_rate    = '~1.5-2%'
trend       = 'Stable/Mature'
notes       = '12,700+ Steam entries; broad tag inflated by blockbusters (Elden Ring Nightreign $130M). Deckbuilder + action subgenres outperform. Needs clear mechanical identity.'

ws.append_row([verdict, genre, solo, team4, hit_rate, trend, notes])
print(f'Appended: {verdict} | {genre} | {hit_rate} | {trend}')

# Export CSV
csv_path = BASE / 'data' / 'genre-viability.csv'
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, 'w', newline='', encoding='utf-8') as f:
    csv.writer(f).writerows(ws.get_all_values())
print(f'Exported: data/genre-viability.csv')

# Git commit + push
def _git(*args):
    r = subprocess.run(['git'] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git('add', 'data/genre-viability.csv')
_git('commit', '-m', f'add: Roguelike to viability ratings - {date.today().isoformat()}')
_git('push', 'origin', 'main')
print('Git: committed and pushed.')
