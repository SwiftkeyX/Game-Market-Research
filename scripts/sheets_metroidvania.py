# -*- coding: utf-8 -*-
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import csv, subprocess, time
from pathlib import Path

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = ["Genre","Game","Year","Revenue","Review","Team size",
           "Trailer","UI","Art style","Feature","Scope","Content","Replayability"]

COLOR_HEADER  = {"red": 0.812, "green": 0.886, "blue": 0.953}
COLOR_HIGH    = {"red": 0.851, "green": 0.918, "blue": 0.827}
COLOR_MID     = {"red": 1.0,   "green": 0.949, "blue": 0.800}
COLOR_FAILURE = {"red": 1.0,   "green": 0.902, "blue": 0.902}
COLOR_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

game_rows = [
    # HIGH TIER
    {
        "genre":"Metroidvania","game":"Hollow Knight: Silksong","tier":"HIGH",
        "year":2025,"revenue":"$100M+","review":"140,381","team_size":"Team of 3 (Team Cherry)",
        "trailer":"3-min cinematic reveal; Hornet movement and boss fights shown in first 10s; hand-animated sequences; orchestral build; immediate genre clarity",
        "ui":"Minimal HUD; clean silk-thread map with gradual reveal; no combat UI clutter; inventory accessed via pause",
        "art_style":"Hand-drawn 2D gothic fantasy, every frame hand-animated, muted jewel tones, intricate enemy silhouettes",
        "feature":"Massive interconnected world; diagonal movement axis; 50+ boss encounters; NPC quest system; silk-weaving traversal; multiple tool types",
        "scope":"~30-50 hrs; larger than Hollow Knight 1; 7+ biomes; full release Sept 2025",
        "content":"Gilded cities, lakes of fire, misted moors; universal Metacritic acclaim; 97% critics recommended; 7M+ copies by Dec 2025; 140K+ reviews",
        "replayability":"High -- completionist collectibles, quest chains, boss challenges, speedrunning. 'Still discovering things at 80 hours' -- Steam",
    },
    {
        "genre":"Metroidvania","game":"Nine Sols","tier":"HIGH",
        "year":2024,"revenue":"~$17-20M (confirmed)","review":"~35,000","team_size":"Team of ~12 (Red Candle Games)",
        "trailer":"High-energy parry combat showcase; deflect mechanic front-and-center in first 5s; manga-panel cutscene style; Dao-punk aesthetic immediately clear",
        "ui":"Clean minimal HUD; manga-style dialogue boxes; clear health/focus bars; no excess panel clutter",
        "art_style":"Hand-drawn anime sprites, manga-inspired cinematic cutscenes, biopunk Taoism aesthetic, 2D side-scrolling",
        "feature":"Sekiro-inspired parry/deflect combat; large connected map ability-gated; equipment upgrades; lore-rich narrative; two endings (True Ending requires full exploration)",
        "scope":"20-30 hrs; 9 major areas; large boss roster; no Early Access",
        "content":"Dao-punk sci-fi blending Taoism + biopunk; 'Best narrative in a metroidvania I have played' -- Steam; 95% positive 35K reviews",
        "replayability":"Moderate -- True Ending requires full exploration; alternate combat styles; high positive rating",
    },
    {
        "genre":"Metroidvania","game":"ANIMAL WELL","tier":"HIGH",
        "year":2024,"revenue":"~$10-21M (confirmed)","review":"~23,938","team_size":"Solo dev (Billy Basso), pub. Bigmode",
        "trailer":"Atmospheric mystery trailer; zero combat shown; surrealist mood-first; no text or tutorial; environmental secrets hinted throughout",
        "ui":"Virtually no HUD; all information embedded in environment; inventory hidden until discovered in-game",
        "art_style":"Pixel art, Commodore 64-inspired palette, modern physics/particle lighting, lush cascading environmental details",
        "feature":"Layered puzzle-mystery exploration; items as environmental keys; 64 secret eggs; no explicit tutorial; custom C++ engine; multiple hidden puzzle layers beneath main game",
        "scope":"~6-10 hrs main campaign; 20+ hrs completionist (all 64 eggs); community ARG puzzle hunting continues post-launch",
        "content":"'A second larger game hidden within the first' -- dev; community solved meta-puzzles months post-launch. 'Still finding new secrets at 100 hours' -- Reddit; 95% positive",
        "replayability":"Very high -- community ARG puzzle hunting, multiple secret layers, speedrunning. Solo dev benchmark for this genre.",
    },
    {
        "genre":"Metroidvania","game":"Blasphemous 2","tier":"HIGH",
        "year":2023,"revenue":"~$20-34M est.","review":"~18,296","team_size":"Team of ~18 (The Game Kitchen)",
        "trailer":"Dark cinematic reveal; boss fight showcase; Spanish Catholic gothic horror aesthetic clear in first 30s; strong music; pixel art quality prominent",
        "ui":"Clean dark-themed inventory overlay; readable health/fervor bars; minimal screen clutter during combat",
        "art_style":"Pixel art, Spanish Catholic gothic horror iconography, fluid hand-crafted animations, grotesque enemy designs",
        "feature":"3 weapons (each unlocks different traversal paths); large connected map; hidden NPC questlines; magic spell system; deep religious lore; DLC expansion",
        "scope":"~20 hrs full exploration; massive map with collectibles; full release Aug 2023",
        "content":"Rich Catholic-horror world; NPCs with full questlines; DLC added. 'Bigger, deeper, better looking than the first'; 90% positive 18K reviews",
        "replayability":"Moderate -- different weapon start changes routing; multiple endings; NG+ mode; DLC content",
    },
    {
        "genre":"Metroidvania","game":"ENDER MAGNOLIA: Bloom in the Mist","tier":"HIGH",
        "year":2025,"revenue":"~$20-33M est.","review":"~17,917","team_size":"Team of 29 (Adglobe + Live Wire, pub. Binary Haze)",
        "trailer":"Anime-quality cinematic; companion Homunculi system highlighted; emotional narrative beats; vibrant environments shown early",
        "ui":"Clean anime-styled inventory; companion ability menu readable; map with progress indicators; minimal HUD during exploration",
        "art_style":"2D anime-style, vibrant dark fantasy, hand-drawn environments, expressive character designs, painterly backgrounds",
        "feature":"30 skills via Homunculi companions; difficulty modes; two endings; dark fantasy narrative; sequel to Ender Lilies (43K reviews)",
        "scope":"8-10 hrs main; 15-20 hrs full; 23+ hrs completionist (all relics/collectibles)",
        "content":"Dark world with rich environmental storytelling; 'An incredible follow-up'; 97-98% positive across ~18K reviews",
        "replayability":"Moderate -- two endings require different routes; full relic/collectible hunting; companion customization",
    },
    {
        "genre":"Metroidvania","game":"Pseudoregalia","tier":"HIGH",
        "year":2023,"revenue":"~$4.2M (confirmed)","review":"~15,678","team_size":"Solo dev (rittzler)",
        "trailer":"Movement showcase first -- somersaults, wall-runs, aerial combos in opening 10s; no explanation; very short trailer; $5.99 price shown",
        "ui":"Minimal 3D sidebar ability list; clean inventory; N64-era aesthetic applied to UI elements",
        "art_style":"Low-poly 3D, N64 5th-gen aesthetic, dreamlike castle environments, muted palette with strong silhouettes",
        "feature":"3D platformer-metroidvania hybrid; open-ended Castle Sansa exploration; movement ability unlocks change traversal; $5.99 price point",
        "scope":"Short -- 3-6 hrs for completion; focus is movement mastery over length",
        "content":"'Best movement in any metroidvania' -- Reddit; 97% positive. Solo dev at $5.99 earning $4.2M gross = strong case for movement-first niche positioning",
        "replayability":"Very high -- movement routing, speedrunning community, alternate ability paths. Self-taught dev sold 200K+ copies.",
    },
    {
        "genre":"Metroidvania","game":"Bo: Path of the Teal Lotus","tier":"HIGH",
        "year":2024,"revenue":"~$2M est.","review":"~1,369","team_size":"Team of 2-4 (Squid Shock Studios, debut)",
        "trailer":"Stylized trailer; Japanese mythology world shown first; bo staff combat and papercraft art prominent; clear aesthetic identity",
        "ui":"Clean minimal HUD; Japanese-aesthetic iconography; Daruma companion display; ability indicators clear",
        "art_style":"Hand-drawn 2D, Japanese mythology-inspired, colorful papercraft aesthetic, traditional brushwork influences, frame-by-frame animation",
        "feature":"Shapeshifting bo staff with tea-based ability transformations; Daruma companion collection; Japanese folklore world; debut studio built in Unity",
        "scope":"~8-12 hrs; moderate map; no DLC at launch",
        "content":"Art praised highly; reception split on depth. GameSpot: 'beautiful but safe -- A Hollow Night.' 85% positive; publisher (Humble Games) shut down mid-dev, shipped anyway",
        "replayability":"Limited -- one playthrough; fewer secrets than genre peers; no NG+",
    },
    # MID TIER
    {
        "genre":"Metroidvania","game":"The Last Case of Benedict Fox","tier":"MID",
        "year":2023,"revenue":"~$1.7M est.","review":"906","team_size":"Team of ~12 (Plot Twist)",
        "trailer":"Gothic noir cinematic; mind-diving mechanic teased; Burton-esque aesthetic clear; mystery tone strong",
        "ui":"Somewhat complex clue/map system; reviewers noted UI friction as a barrier",
        "art_style":"Tim Burton-esque gothic noir, dark atmospheric 2D, hand-animated character sprites",
        "feature":"Lovecraftian murder mystery + metroidvania; mind-diving into dead people mechanic; elaborate puzzles; jazz-noir soundtrack; Xbox-backed publishing",
        "scope":"~6-8 hrs; linear mystery progression",
        "content":"'Gameplay falls apart where it should shine' -- PC Gamer; Metascore 73; 71% positive (Mixed). Aesthetic strong, combat/UX weak. Cautionary tale: great concept, poor execution",
        "replayability":"Low -- linear mystery; little replay incentive",
    },
    {
        "genre":"Metroidvania","game":"Tales of Kenzera: ZAU","tier":"MID",
        "year":2024,"revenue":"~$1.3M est.","review":"864","team_size":"Team of ~30 (Surgent Studios / EA Originals)",
        "trailer":"Emotional story-first reveal; father-son grief narrative; Bantu culture aesthetic shown; dual mask system introduced",
        "ui":"Dual mask HUD (sun/moon) clearly displayed; clean ability indicators",
        "art_style":"Vibrant 2D, Bantu culture-inspired African mythology, warm color palette, fluid animations",
        "feature":"Sun mask (melee) + moon mask (ranged) dual system; shaman powers; African mythology world; ~8 hrs; personal grief narrative by founder Abubakar Salim",
        "scope":"~8 hrs; moderate map; published by EA Originals",
        "content":"'Spirit over substance' -- Game Informer. Personal grief narrative praised; combat thin. Surgent faced layoffs post-launch. EA publishing + story ambition + thin combat = Mixed commercial outcome",
        "replayability":"Low -- linear story; few progression branches",
    },
    {
        "genre":"Metroidvania","game":"Arzette: The Jewel of Faramore","tier":"MID",
        "year":2024,"revenue":"~$1.3M est.","review":"848","team_size":"Small team (Seth Fulkerson + collaborators)",
        "trailer":"Deliberate CD-i callback -- campy voice acting, colorful cartoon world, hand-painted backgrounds; niche audience immediately recognized",
        "ui":"Clean retro-styled 2D platformer HUD; minimal",
        "art_style":"CD-i animated cartoon style, hand-painted backgrounds, fully voice-acted campy cinematics, colorful cast",
        "feature":"Spiritual successor to CD-i Zelda games; diverse locations with secrets; colorful cast; deliberate retro FMV aesthetic",
        "scope":"~5-8 hrs; moderate world size",
        "content":"96% positive but very small addressable market. 'Exactly what it promises' -- enthusiast fans. Niche positioning caps revenue despite critical love",
        "replayability":"Low -- completionist secrets but no major branching",
    },
    {
        "genre":"Metroidvania","game":"Ultros","tier":"MID",
        "year":2024,"revenue":"~$1.1M est.","review":"723","team_size":"Team of ~5 (Hadoque, Sweden)",
        "trailer":"Psychedelic visual showcase; cosmic horror atmosphere; El Huervo art style front and center; unusual aesthetic signals niche audience",
        "ui":"Unusual organic interface matching psychedelic art; some readability friction noted by reviewers",
        "art_style":"Psychedelic 2D, cosmic uterus aesthetic, vibrant alien art by El Huervo (Hotline Miami artist), highly stylized",
        "feature":"Time-loop mechanic; plant-growing progression system; alien cosmic world; non-combat progression options; artistic collaboration with El Huervo",
        "scope":"~10-15 hrs; rich explorable world",
        "content":"'A drop-dead gorgeous head trip' -- GamesRadar; 79% positive. Art-first at cost of gameplay clarity. Strong visual identity, weak gameplay loop cohesion",
        "replayability":"Moderate -- time-loop discovery, cosmic secrets; constrained by audience ceiling",
    },
    {
        "genre":"Metroidvania","game":"Shattered Divinities","tier":"MID",
        "year":2025,"revenue":"~$450K est.","review":"~403","team_size":"Unknown small indie team",
        "trailer":"Fast-paced action showcase across nine divine realms; combat abilities highlighted; pixel art quality shown",
        "ui":"Clean and functional pixel UI; ability indicators clear",
        "art_style":"2D pixel art, nine realms each with distinct visual palette, varied environmental styles",
        "feature":"9 broken divine realms; 100+ intricate maps; 20 bosses; Phase Step, Ethereal Step, Bulwark, Assault, Invocation ability set",
        "scope":"Medium -- 9 biomes; 100+ maps; moderate campaign length",
        "content":"'Low cost, well done -- controls are tight, upgrades constant' -- Steam; 93% positive. Good game that stayed under the radar in a crowded genre window",
        "replayability":"Moderate -- boss rush, full map completion; limited build variety",
    },
    {
        "genre":"Metroidvania","game":"DOOMBLADE","tier":"MID",
        "year":2023,"revenue":"~$400K est.","review":"~356","team_size":"Small team (Muro Studios)",
        "trailer":"Combat-focused showcase; gritty dark aesthetic; vengeful quest tone; roguelike unpredictability suggested",
        "ui":"Standard platformer HUD; minimal; readable",
        "art_style":"Gritty dark 2D, detailed atmospheric environments, horror-dread aesthetic, high-contrast enemy designs",
        "feature":"Tight melee combat + roguelike elements; unpredictable run composition; vengeful quest narrative against Dread Lords",
        "scope":"Medium length with roguelike replay structure",
        "content":"81% positive; criticized for unconventional lock-on and contact damage mid-flight. Niche audience ceiling. Punishing in the wrong ways per some reviewers",
        "replayability":"Moderate -- roguelike randomization; constrained by small discovery audience",
    },
    {
        "genre":"Metroidvania","game":"Curse of the Sea Rats","tier":"MID",
        "year":2023,"revenue":"~$300K est.","review":"198","team_size":"Small team (Petoons Studio, Barcelona)",
        "trailer":"Hand-drawn art showcased prominently; 4-player co-op angle shown; 'ratoidvania' branding; colorful tone",
        "ui":"4-player capable HUD; can feel cluttered in co-op; readable solo",
        "art_style":"Hand-drawn 2D, 4 animal protagonists, Spanish studio animation quality, colorful warm palette",
        "feature":"4 playable characters with unique skill sets; 4-player co-op support; non-linear interconnected world; boss encounters tied to character abilities",
        "scope":"~8-12 hrs solo; moderate map",
        "content":"Praised for art quality; reception mixed on solo depth. 'Feels thin without friends' -- Steam. Co-op hook underpowered for the promise made",
        "replayability":"Moderate with friends; low solo",
    },
    # FAILURE TIER
    {
        "genre":"Metroidvania","game":"MARS 2120","tier":"FAILURE",
        "year":2024,"revenue":"~$35K est.","review":"47","team_size":"Small team (QUByte Interactive, Brazil)",
        "trailer":"Generic sci-fi environments; combat demonstrated without polish; no hook or identity established",
        "ui":"Cluttered; input registration issues reported; no tutorial for controls",
        "art_style":"3D sci-fi, low-budget production values, generic visual identity, flat environments",
        "feature":"Sci-fi metroidvania; ranged + melee combat; boss encounters; no tutorial system",
        "scope":"~4-6 hrs; short campaign",
        "content":"'Janky experience made for devoted genre fans only' -- reviewer. No tutorial, weak ranged attacks, overpowered bosses. 80% positive (47 reviews) = near-invisible on Steam. Xbox port cancelled.",
        "replayability":"Very low -- poor gameplay loop; no NG+ or extra modes",
    },
    {
        "genre":"Metroidvania","game":"Clockwork Ambrosia","tier":"FAILURE",
        "year":2025,"revenue":"~$80K est.","review":"71","team_size":"Unknown (solo or very small team)",
        "trailer":"Steampunk exploration and combat shown; calculated arsenal premise highlighted; modest production value",
        "ui":"Standard pixel art HUD; functional; no readability issues",
        "art_style":"Pixel art steampunk, clockwork world environments, mechanical enemy designs",
        "feature":"Arsenal-based combat system; steampunk metroidvania exploration; map progression gating",
        "scope":"Short to medium campaign",
        "content":"83% positive but invisible on Steam. PC Gamer cited as 'quality game few have heard of.' Launched 2025 into dead genre chart window -- no metroidvania penetrated Steam weekly charts except Silksong.",
        "replayability":"Low",
    },
    {
        "genre":"Metroidvania","game":"Narvas","tier":"FAILURE",
        "year":2022,"revenue":"~$0 (given free Mar 2026)","review":"~50","team_size":"Unknown developer",
        "trailer":"Generic sci-fi metroidvania showcase; no strong identity signal; no hook",
        "ui":"Standard HUD; functional",
        "art_style":"Sci-fi 2D, stylized but generic, muted palette",
        "feature":"Sci-fi world exploration; metroidvania ability gating; standard combat",
        "scope":"Short",
        "content":"62% positive (Mixed); given away free on Steam in March 2026 -- near-zero commercial value. 'Fine but forgettable.' Baseline genre failure: competent but no identity or discovery.",
        "replayability":"Very low",
    },
]

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

today = date.today().isoformat()
genre_display = "Metroidvania"
tab_name = f"Metroidvania"

MAX_ARCHIVES = 3

# Archive existing tab
try:
    old_ws = sh.worksheet(tab_name)
    old_vals = old_ws.get_all_values()
    if len(old_vals) > 1:
        archive_name = f"Metroidvania ({today})"
        old_ws.update_title(archive_name)
    else:
        sh.del_worksheet(old_ws)
except gspread.WorksheetNotFound:
    pass

# Delete oldest archives beyond rolling limit
archive_prefix = f"Metroidvania ("
all_tabs = sh.worksheets()
archives = [wks for wks in all_tabs if wks.title.startswith(archive_prefix)]
archives.sort(key=lambda wks: wks.title)
while len(archives) > MAX_ARCHIVES:
    sh.del_worksheet(archives.pop(0))

ws = sh.add_worksheet(title=tab_name, rows=500, cols=len(HEADERS) + 2)
ws.append_row(HEADERS)
ws.format(f"A1:{chr(64 + len(HEADERS))}1", {
    "backgroundColor": COLOR_HEADER,
    "textFormat": {"bold": True}
})

_col_widths = [160, 200, 60, 120, 80, 120, 200, 180, 160, 350, 250, 300, 300]
sh.batch_update({"requests": [
    {
        "updateDimensionProperties": {
            "range": {"sheetId": ws.id, "dimension": "COLUMNS",
                      "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }
    for i, px in enumerate(_col_widths)
] + [
    {
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 1,
                "startColumnIndex": 6,
                "endColumnIndex": 13,
            },
            "cell": {"userEnteredFormat": {
                "wrapStrategy": "WRAP",
                "verticalAlignment": "TOP",
            }},
            "fields": "userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment",
        }
    },
    {
        "updateSheetProperties": {
            "properties": {
                "sheetId": ws.id,
                "gridProperties": {"frozenRowCount": 1},
            },
            "fields": "gridProperties.frozenRowCount",
        }
    },
]})

sheet_rows = []
tiers = []
for g in game_rows:
    sheet_rows.append([
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ])
    tiers.append(g["tier"])

ws.append_rows(sheet_rows, value_input_option="USER_ENTERED")

# Auto-resize rows
sh.batch_update({"requests": [{
    "autoResizeDimensions": {
        "dimensions": {
            "sheetId": ws.id,
            "dimension": "ROWS",
            "startIndex": 1,
        }
    }
}]})

# Color rows by tier
all_vals = ws.get_all_values()
total_rows = len(all_vals)
start_row = total_rows - len(sheet_rows) + 1

format_requests = []
for i, (row_data, tier) in enumerate(zip(sheet_rows, tiers), start=start_row):
    color = {"HIGH": COLOR_HIGH, "MID": COLOR_MID, "FAILURE": COLOR_FAILURE}.get(tier, COLOR_WHITE)
    format_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": i - 1,
                "endRowIndex": i,
                "startColumnIndex": 0,
                "endColumnIndex": len(HEADERS),
            },
            "cell": {"userEnteredFormat": {"backgroundColor": color}},
            "fields": "userEnteredFormat.backgroundColor",
        }
    })
    if len(format_requests) >= 20:
        sh.batch_update({"requests": format_requests})
        format_requests = []
        time.sleep(1.2)

if format_requests:
    sh.batch_update({"requests": format_requests})

print(f"Written {len(sheet_rows)} games to tab '{tab_name}'.")

# Export CSV + Git commit
BASE = Path(r"C:\Organized Files\My Game Asset\Game-Research")
data_dir = BASE / "data" / "competitors"
data_dir.mkdir(parents=True, exist_ok=True)
csv_path = data_dir / "metroidvania.csv"
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ws.get_all_values())
print(f"Exported: {csv_path.relative_to(BASE)}")

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

_git("add", "data/competitors/metroidvania.csv", "research-log.md")
_git("commit", "-m", f"research: Metroidvania competitor data - {today}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
