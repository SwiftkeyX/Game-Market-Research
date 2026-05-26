import gspread
from google.oauth2.service_account import Credentials
import sys
sys.stdout.reconfigure(encoding='utf-8')

CREDENTIALS_FILE = r'C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json'
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh = client.open_by_url(SHEET_URL)

HEADERS = ['Verdict','Genre','Solo Target','Team of 4','Hit Rate','Trend','Notes']

GENRES = [
  ['GO','Roguelike Deckbuilder','$100K-$500K','$500K-$2M','5.1%','Growing','3 hits Q1 2026 (normally 1/yr). Best solo genre. 200-500 cards manageable.'],
  ['GO','Idle / Incremental','$100K-$300K','$200K-$600K','~3% growing','Golden age','3 hits 2022 -> 16 in 2024 -> 12 in Q2 2025. 2-3 month solo build. Low art bar.'],
  ['GO','Job Simulation 3D','$150K-$400K','$500K-$2M','4.1%','Overtaking Mgmt','Blue-collar 3D first-person dominates. Hard solo without 3D skills.'],
  ['GO','Horror (any combo)','$100K-$400K','$300K-$1.5M','3.2%','Cooling Q1 2026','Flexible: Horror+Casino, Horror+Idle, Horror+Co-op all work.'],
  ['GO','Friend-Slop Co-op','$100K-$500K','$300K-$1.5M','Top Q1 2026','Rising','Most hits Q1 2026. Build in 4-8 weeks. Viral streaming potential.'],
  ['GO','Narrative Story-Driven','$100K-$500K','$300K-$1.5M','51 hits 2025','Strongly rising','#1 genre by hit count 2025. Story + light gameplay systems.'],
  ['GO','Pixel Art / Indie RPG','$100K-$400K','$300K-$1.5M','~2.4% stable','Consistent','18 hits Q2 2025. Top-down, 8-15hrs focused scope.'],
  ['GO','Colony Sim / Factory','N/A (solo skip)','$500K-$3M','6.4% hybrid','Rising','48% of top Next Fest demos. Team only. Expect 2-4yr dev time.'],
  ['GO','Farming / Life Sim','$100K-$500K','$300K-$1.5M','Consistent hits','Stable','Multiple hits annually. Need hook beyond just farming.'],
  ['GO','Auto-battler','$100K-$400K','$300K-$1.5M','Phase 3 emerging','Emerging','Low saturation window open. Low art burden, high systems depth.'],
  ['CAUTION','Tower Defense Hybrid','$100K-$400K','$300K-$1.5M','Reviving','Rising (needs meta)','5 hits Q2 2025 vs 8 all 2024. Meta layer non-negotiable.'],
  ['CAUTION','Tactical Strategy RPG','$100K-$500K','$400K-$2M','~3-4%','Steady','Design-heavy not art-heavy. Scope balloons fast; strict discipline needed.'],
  ['CAUTION','Open World Survival Craft','N/A (solo skip)','$500K-$3M','20.8%','Dropped from top 5','Highest hit rate but scope brutal (2-4 years). Team only.'],
  ['CAUTION','Action Roguelite','$150K-$500K','$500K-$2M','~2-3%','Stable/crowded','Dead Cells + Hades set punishing quality bar. Hard solo without art skills.'],
  ['CAUTION','Cozy Management','$100K-$400K','$300K-$1.2M','Moderate','Stable','Management-cozy hits; adventure-cozy does not. Build loop first.'],
  ['CAUTION','Metroidvania','$150K-$600K','$400K-$2M','~4% / 0 Q1 hits','Declining freq.','Zero hits Q1 2025. Hollow Knight bar brutal. Must be great, not just good.'],
  ['CAUTION','Visual Novel (pure)','$50K-$200K','$100K-$500K','Skewed (Asian pub)','Growing (skewed)','51 hits mostly Chinese FMV. Western pure VN rate far lower than aggregate.'],
  ['CAUTION','JRPG','$100K-$400K','$300K-$1.2M','Gatekept','Stable','Dominated by Atlus/Falcom back-catalog. Budget 2-3 years for modest scope.'],
  ['CAUTION','3D Platformer','$80K-$300K','$200K-$800K','1.46%','No hits Q1 2025','3D physics/camera exponentially harder. AAA polish expected.'],
  ['CAUTION','Rage / Precision Platformer','$50K-$200K','$100K-$400K','Phase 4','Declining','Window closing since 2022. Viral streaming hook essential.'],
  ['AVOID','2D Platformer','N/A','N/A','0.18%','Zero hits Q1 2026','Near-zero. Free browser games dominate.'],
  ['AVOID','Point-and-Click Adventure','N/A','N/A','0.18%','Declining','LucasArts quality expected. Near-zero modern indie revenue.'],
  ['AVOID','Multiplayer Shooter','N/A','N/A','Dies post-launch','Collapses fast','Steel Hunters: 4473 -> 45 players. Cannot sustain playerbase.'],
  ['AVOID','Pure Puzzle','N/A','N/A','0.34%','No improvement','As primary tag fails. Players get free content on mobile.'],
  ['AVOID','Bullet Heaven Clone','N/A','N/A','<1%','Saturated 2023','Only 1 hit in all of 2024. Window closed.'],
  ['AVOID','VR Game','N/A','N/A','0 hits Q1 2026','Declining','VR has no audience at scale. Hardware penetration too low.'],
  ['AVOID','4X Strategy','N/A','N/A','Near zero','No improvement','Civ/Stellaris dominate. 3-5yr dev minimum.'],
  ['AVOID','Racing Game','N/A','N/A','Near zero','AAA dominated','Forza/GT capture all intent. Physics judged harshly.'],
  ['AVOID','Fighting Game','N/A','N/A','Near zero','Gatekept','Capcom/ArcSys dominate. Balancing + rollback netcode required.'],
  ['AVOID','Music / Rhythm','N/A','N/A','Near zero paid','F2P kills it','Osu free, Muse Dash F2P. No paid market without compelling IP hook.'],
]

# Create ratings tab
try:
    ws = sh.worksheet('Genre Viability Ratings (GO / CAUTION / AVOID)')
    ws.clear()
    print('Cleared existing tab.')
except Exception:
    ws = sh.add_worksheet('Genre Viability Ratings (GO / CAUTION / AVOID)', rows=100, cols=10)
    print('Created new ratings tab.')

ws.append_row(HEADERS)
ws.append_rows(GENRES)
print(f'Written {len(GENRES)} genres to ratings tab.')

# Create history snapshot
try:
    hist = sh.worksheet('v1 - May 2026')
    hist.clear()
except Exception:
    hist = sh.add_worksheet('v1 - May 2026', rows=100, cols=10)
hist.append_row(['SNAPSHOT v1 - May 2026 | READ ONLY | Tier movements: None (first version)'])
hist.append_row(HEADERS)
hist.append_rows(GENRES)
print('History snapshot written to v1 - May 2026.')

# Check which of the 3 requested genres already exist
existing_lower = [r[1].lower() for r in GENRES]
for check in ['roguelike', 'autobattler', 'auto-battler', 'metroidvania']:
    print(f'CHECK {check}: {check in existing_lower}')
