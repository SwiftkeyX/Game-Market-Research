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

# --- Step 1: Delete the 5 broad genres ---
TO_DELETE = [
    'Horror (any combo)',
    'Pixel Art / Indie RPG',
    'Colony Sim / Factory',
    'Narrative Story-Driven',
    'Multiplayer Shooter',
]

for genre_name in TO_DELETE:
    all_vals = ws.get_all_values()
    for i, row in enumerate(all_vals):
        if row and row[1] == genre_name:
            ws.delete_rows(i + 1)  # 1-indexed
            print(f'Deleted: {genre_name} (row {i+1})')
            break
    else:
        print(f'Not found (skip): {genre_name}')

# --- Step 2: Append 11 replacement rows ---
NEW_ROWS = [
    # Horror (any combo) -> 3 specific entries
    ['GO',      'Horror Co-op',                   'N/A (needs server)', '$200K-$1M',     '~3-4%',        'Rising',           'Half of all friend-slop hits are horror. Phasmophobia/Lethal Company set blueprint. Viral streaming essential. 3-4 week build possible.'],
    ['CAUTION', 'Horror Puzzle / Exploration',     '$80K-$300K',         '$250K-$1.2M',   '~2-3%',        'Cooling',          'Horror #1 in top sellers 3 years running but cooling — 7 hits 2025 vs 3 Q1 2026. Solo-friendly: atmosphere over mechanics. Amnesia/SIGNALIS bar is high.'],
    ['CAUTION', 'Horror Idle',                     '$80K-$250K',         '$150K-$600K',   '~1.5-2%',      'Emerging',         'Horror x Idle crossover is a proven combo but thin data. Lower hit rate than pure idle (~3%). Small underserved niche; 2-3 month build.'],

    # Pixel Art / Indie RPG -> 1 renamed entry
    ['GO',      'Top-Down Pixel RPG',              '$100K-$400K',        '$300K-$1.5M',   '~2.4%',        'Consistent',       '18 hits Q2 2025. Top-down perspective, 8-15hrs focused story. RPG Maker viable. Anime art increasingly dominant (ties to narrative boom).'],

    # Colony Sim / Factory -> 2 specific entries
    ['GO',      'Colony Sim',                      'N/A (solo skip)',    '$500K-$3M',     '~4-5%',        'Rising',           'RimWorld/Dwarf Fortress tier has devoted community. AI behavior + simulation systems essential. 2-4yr dev minimum. Solo effectively impossible.'],
    ['CAUTION', 'Factory / Automation',            'N/A (solo skip)',    '$400K-$2M',     '~3-4%',        'Stable',           'Factorio dominates mindshare. New entries need a distinct twist (space layer, survival). 2+ yr dev minimum. Systems complexity punishing solo.'],

    # Narrative Story-Driven -> 2 specific entries
    ['GO',      'Narrative RPG / Choice Game',     '$100K-$500K',        '$300K-$1.5M',   'High (~51 hits 2025)', 'Strongly rising', '#1 genre by hit count 2025. Anime art dominant (Chinese market). Western narrative RPG viable. Text-heavy = low art bar. Choice architecture required.'],
    ['AVOID',   'Walking Sim / Environmental Narrative', 'N/A',          'N/A',           '<0.5%',        'Declining',        'Genre peaked 2016-2018; hit rate near zero since. Audience moved to narrative RPGs and VNs. Only viable with exceptional world/concept (Firewatch-tier).'],

    # Multiplayer Shooter -> 3 specific entries
    ['AVOID',   'Battle Royale',                   'N/A',                'N/A',           'Near zero',    'Collapsing',       'Playtime dropped -27% in 2025. Fortnite/PUBG/Warzone dominate. Highguard shut down in 45 days. Live-service requirement makes indie entry impossible.'],
    ['AVOID',   'Hero Shooter',                    'N/A',                'N/A',           'Near zero',    'Saturated',        'Overwatch/Valorant dominate. Live-service fatigue widespread. Highguard failed in 45 days 2026. Cannot sustain playerbase without massive marketing budget.'],
    ['CAUTION', 'Extraction Shooter',              'N/A (solo skip)',    '$300K-$1.5M',   '~0.5-1%',      'Niche/Stable',     '223 Steam entries; Arc Raiders ($252M) skews category data — ex-EA team. True indie hit rate low. Requires server infra + anticheat. Not solo viable.'],
]

ws.append_rows(NEW_ROWS)
print(f'\nAppended {len(NEW_ROWS)} new rows.')

# --- Step 3: Export CSV ---
csv_path = BASE / 'data' / 'genre-viability.csv'
csv_path.parent.mkdir(parents=True, exist_ok=True)
with open(csv_path, 'w', newline='', encoding='utf-8') as f:
    csv.writer(f).writerows(ws.get_all_values())
print(f'Exported: data/genre-viability.csv')

# --- Step 4: Git commit + push ---
def _git(*args):
    r = subprocess.run(['git'] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git('add', 'data/genre-viability.csv')
_git('commit', '-m', f'refactor: split 5 broad genres into 11 specific subgenres - {date.today().isoformat()}')
_git('push', 'origin', 'main')
print('Git: committed and pushed.')

# --- Summary ---
rows = ws.get_all_records()
print(f'\nSheet now has {len(rows)} genres.')
