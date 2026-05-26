"""
2D Platformer competitor data - Google Sheets upload script
Writes to tab: 2D Platformer
Exports CSV: data/competitors/2d-platformer.csv
Git: commit + push
"""
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import time
import csv
import subprocess
from pathlib import Path

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADERS = [
    "Genre", "Game", "Year", "Revenue", "Review",
    "Team size", "Trailer", "UI", "Art style",
    "Feature", "Scope", "Content", "Replayability",
]

COLOR_HEADER  = {"red": 0.812, "green": 0.886, "blue": 0.953}
COLOR_HIGH    = {"red": 0.851, "green": 0.918, "blue": 0.827}
COLOR_MID     = {"red": 1.0,   "green": 0.949, "blue": 0.800}
COLOR_FAILURE = {"red": 1.0,   "green": 0.902, "blue": 0.902}
COLOR_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

today      = date.today().isoformat()
tab_name   = "2D Platformer"
MAX_ARCHIVES = 3

# Archive existing tab
try:
    old_ws = sh.worksheet(tab_name)
    old_vals = old_ws.get_all_values()
    if len(old_vals) > 1:
        archive_name = f"2D Platformer ({today})"
        old_ws.update_title(archive_name)
        print(f"Archived existing tab as: {archive_name}")
    else:
        sh.del_worksheet(old_ws)
except gspread.WorksheetNotFound:
    pass

# Trim old archives
archive_prefix = "2D Platformer ("
all_tabs = sh.worksheets()
archives = [ws for ws in all_tabs if ws.title.startswith(archive_prefix)]
archives.sort(key=lambda ws: ws.title)
while len(archives) > MAX_ARCHIVES:
    sh.del_worksheet(archives.pop(0))

ws = sh.add_worksheet(title=tab_name, rows=500, cols=len(HEADERS) + 2)
ws.append_row(HEADERS)
ws.format(f"A1:{chr(64 + len(HEADERS))}1", {
    "backgroundColor": COLOR_HEADER,
    "textFormat": {"bold": True}
})

_col_widths = [160, 200, 60, 160, 80, 160, 200, 180, 160, 350, 250, 300, 300]
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

game_rows = [
    # ---- HIGH TIER ----
    {
        "genre": "2D Platformer", "game": "Pizza Tower", "tier": "HIGH", "year": 2023,
        "revenue": "~$72M+ est.", "review": "71,408",
        "team_size": "Solo (Tour De Pizza / McPig)",
        "trailer": "Multiple viral trailers showing frantic Wario Land-speed runs; meme-generating clips drove massive pre-launch buzz; shows momentum system in first 5s",
        "ui": "Minimal HUD - speed meter and Peppino state visible; no clutter; arcade score display clean and readable",
        "art_style": "High-res pixel art, 1990s Nicktoon-inspired, vibrant chaotic palette, exaggerated rubber-hose animations",
        "feature": "- Momentum-based speed platforming (walk to sprint to chaos)\n- 5 floors x multiple levels; secret bonus rooms in every level\n- Score rank system (D to P); taunts + parries\n- 3 campaigns; 2 playable characters; 15+ boss fights",
        "scope": "5-6hr first playthrough; 15+ levels; P-rank chase adds 50-100+ hrs; launched full Jan 2023",
        "content": "5 floors with unique biomes; cheese dragons; secret bonus dungeons per level. 'More content than games 3x the price' - Steam",
        "replayability": "Extremely high - P-rank pursuit drives hundreds of hours; daily challenge community; active mod scene. '400+ hrs chasing P ranks' - Reddit",
    },
    {
        "genre": "2D Platformer", "game": "SANABI", "tier": "HIGH", "year": 2023,
        "revenue": "~$20-30M est.", "review": "43,920",
        "team_size": "Team of 5 (WonderPotion) + NEOWIZ pub.",
        "trailer": "Shows grappling hook mechanic clearly with stylish cyberpunk action; music-driven pacing; introduces story stakes early",
        "ui": "Clean minimal cyberpunk HUD; health and energy bars readable at glance; boss health bars stylized but clear",
        "art_style": "Pixel art cyberpunk dystopia; neon-lit dark city aesthetic; cel-shaded sprites; vivid neon accents on dark palette",
        "feature": "- Prosthetic grappling arm as sole traversal tool (swing, boost, slam)\n- Combined traversal + weapon in one mechanic\n- Story-driven found-family narrative\n- Diverse boss fights integrating hook\n- Multiple districts with unique visual identity",
        "scope": "8-10hr campaign; multiple chapters; NG+ and boss rush; EA June 2022, full Nov 2023",
        "content": "Multiple city districts; varied enemy types; memorable boss designs; cinematic story. 'Cried at the ending' - Steam",
        "replayability": "Moderate-high; NG+ mode; speedrun community; 'Replayed twice for story details I missed' - Reddit",
    },
    {
        "genre": "2D Platformer", "game": "Neon White", "tier": "HIGH", "year": 2022,
        "revenue": "$8.8M (reported)", "review": "19,085",
        "team_size": "Team of ~6 (Angel Matrix) + Annapurna pub.",
        "trailer": "Fast-cut speedrun showcase with anime aesthetic; card-discarding mechanic shown clearly in first 15s; viral among speedrun and anime communities",
        "ui": "Extremely clean - medal timer, card slots, demon count; highly readable at speed; leaderboard ghost visible",
        "art_style": "Anime 2D character portraits + vibrant neon 3D environments; cel-shaded; bold color-coded card system",
        "feature": "- Card-based abilities: discard cards for special moves (dash, double-jump, rocket)\n- Speedrun-focused 100+ level design\n- Leaderboards + medal grades (Bronze-Ace) baked in\n- Gift/relationship system unlocks story scenes\n- 5 chapters + post-game challenge missions",
        "scope": "6-10hr main story; 100+ levels; post-game challenge levels; significant speedrunning depth",
        "content": "100+ levels; gift items unlock story scenes; challenge levels add 4-6hr post-game. 'Perfect just-one-more-attempt loop' - Steam",
        "replayability": "Very high; friend leaderboard competition; gold/Ace runs add 50+ hrs. 'Replayed every level 200+ times chasing Ace' - speedrun community",
    },
    {
        "genre": "2D Platformer", "game": "Anomaly Agent", "tier": "HIGH", "year": 2024,
        "revenue": "~$10M est.", "review": "8,973",
        "team_size": "Team of ~5 (Phew Phew Games) + GameDev.ist",
        "trailer": "High-energy showcase of cyberpunk beat-em-up combat; stylish anime action cuts; communicates tone immediately",
        "ui": "Clean manga-panel HUD; health/energy minimal but visible; comic book presentation consistent with design",
        "art_style": "Pixel art cyberpunk manga; B&W manga cutscenes; vibrant neon action palette; punchy particle effects",
        "feature": "- Beat-em-up platformer hybrid; agency operative setting\n- Multiple weapons with unique feel\n- Time-stop anomaly mechanic for crowd control\n- Parry system with skill expression\n- Multiple mission levels with varied objectives\n- 3 difficulty modes",
        "scope": "6-8hr campaign; multiple mission levels; 3 difficulties; PS Vita-era action pacing",
        "content": "Multiple agency missions; diverse enemy factions and bosses; operative narrative. 'Best action platformer in years' - Steam",
        "replayability": "Moderate-high; challenge modes; S-rank chase; 'Goes so fast you replay for pure feel' - Reddit",
    },
    {
        "genre": "2D Platformer", "game": "Lil Gator Game", "tier": "HIGH", "year": 2022,
        "revenue": "~$5.9M est.", "review": "3,929",
        "team_size": "Team of 3 (MegaWobble) + Playtonic Friends pub.",
        "trailer": "Adorable wholesome tone; cozy exploration emphasized; slow-paced warmth over action; appeals to cozy audience clearly",
        "ui": "Clean minimal; companion tracking and quest markers subtle; designed for cozy not stressful experience",
        "art_style": "Colorful low-poly 3D; bright cheerful palette; cel-shaded chibi characters; toybox aesthetic with cardboard props",
        "feature": "- Cozy 3D collectathon platformer; no combat\n- Cardboard shield crafting mechanic\n- NPC friendship quests drive progression\n- Sibling relationship narrative\n- Open world exploration with collectibles\n- DLC 'In the Dark' (2025)",
        "scope": "4-6hr main experience; dense NPC and collectible world; DLC adds 2-3hr",
        "content": "Multiple themed island areas; 50+ NPCs with dialogue; heartwarming sibling story. 'Made me cry before credits' - Steam",
        "replayability": "Low-moderate; exploration replay; DLC content; 'Perfect weekend game, played twice' - Reddit",
    },
    {
        "genre": "2D Platformer", "game": "Gravity Circuit", "tier": "HIGH", "year": 2023,
        "revenue": "~$3.8M est.", "review": "3,381",
        "team_size": "Solo (Antti / Domesticated Ant Games) + PID Games pub.",
        "trailer": "Clean showcase of Mega Man-inspired stage progression and robot combat; communicates genre clearly to target audience",
        "ui": "Classic retro HUD - health bar, lives, weapon/energy; faithful to Mega Man inspiration; clean and functional",
        "art_style": "Detailed 2D pixel art; vibrant robot designs; Mega Man Zero/ZX aesthetic; clean readable enemy designs",
        "feature": "- Mega Man-inspired nonlinear stage progression\n- Gundam-inspired robot visual design\n- Fighting-game style combo system\n- Robot companion assists; upgradeable abilities\n- 8 main stages + post-game challenges\n- Boss rush mode",
        "scope": "8-12hr campaign; 8+ robot masters; upgradeable abilities; score attack; ~10 years solo development",
        "content": "8 themed stages; diverse upgrades from bosses; arena challenge mode. 'Best Mega Man since the originals' - Metacritic users",
        "replayability": "High; S-rank chase; boss rush; 'I speedrun this now' - Reddit",
    },
    {
        "genre": "2D Platformer", "game": "ANTONBLAST", "tier": "HIGH", "year": 2024,
        "revenue": "~$4.3M est. (recouped <1 month post-launch)", "review": "2,835",
        "team_size": "Team of ~8-10 (Summitsphere) + Joystick Ventures pub.",
        "trailer": "Explosive high-energy; Wario Land-style destruction shown clearly; viral social media pre-launch; chaotic energy communicated immediately",
        "ui": "Bold arcade-style HUD; Blast Meter visible; collectible tracking minimal; clean for high-speed play",
        "art_style": "Vibrant 2D pixel art; 1990s Nickelodeon cartoon energy; explosive particle effects; bold saturated palette",
        "feature": "- Wario Land-inspired environment destruction mechanics\n- Unique Blast Meter system\n- 12 side-scrolling levels across varied worlds\n- Boss rush + Lime Trials time challenge mode\n- Hard mode; Stage Rush speedrun mode\n- 2 playable characters (Anton, Annie)",
        "scope": "6-8hr campaign; 12 levels; boss rush; EA phase ~6 months; full Dec 2024 with bonus modes",
        "content": "12 levels from Boiler City to Hell; 4 post-launch bosses; diverse destruction physics. 'A love letter to Wario Land' - Metacritic",
        "replayability": "High; speedrun-friendly; Lime Trials replay; 'I play this on lunch breaks still' - Reddit",
    },

    # ---- MID TIER ----
    {
        "genre": "2D Platformer", "game": "Frogun", "tier": "MID", "year": 2022,
        "revenue": "~$570K est.", "review": "507",
        "team_size": "Solo (Marco Venturini / Mokaloca)",
        "trailer": "Charming retro aesthetic; shows grapple-frog mechanic clearly; PS1-era nostalgia evident in visual and sound",
        "ui": "Minimal retro-styled HUD; life/grapple-ammo counter visible; PS1-era presentation",
        "art_style": "Low-poly 3D N64/PS1 era; chunky character models; warm color palette; retro cartridge aesthetic",
        "feature": "- 'Frogun' grappling hook as primary traversal\n- 3D platformer with collecting mechanics\n- Multiple themed worlds\n- Boss fights integrating grapple\n- DLC 'Frogun Encore' (2023) adds content",
        "scope": "4-6hr main campaign; multiple worlds; Encore DLC adds 3-4hr; solid solo dev post-launch support",
        "content": "Multiple themed retro worlds; boss encounters; collectibles. 'Perfectly captures PS1 charm' - Steam",
        "replayability": "Moderate; time trials; DLC adds replay; 'Short but delightful' - Steam",
    },
    {
        "genre": "2D Platformer", "game": "BioGun", "tier": "MID", "year": 2024,
        "revenue": "~$130K (reported estimate)", "review": "477",
        "team_size": "Duo (Dapper Dog Digital, Dallas TX)",
        "trailer": "Colorful inside-body sci-fi aesthetic; Metroidvania loops shown; biome exploration emphasized; niche but clear pitch",
        "ui": "Standard Metroidvania HUD; map + health + weapon; somewhat busy but functional",
        "art_style": "Vibrant 2D sci-fi cartoon; inside-human-body aesthetic; colorful cell/bacteria enemy designs",
        "feature": "- Metroidvania + platformer hybrid; vaccine-themed combat\n- Twin-stick shooter elements\n- Unlockable abilities gate exploration\n- Interconnected body biomes\n- Full map; multiple upgrade paths",
        "scope": "10-15hr Metroidvania; multiple biomes; EA for ~2 years before 2024 full launch",
        "content": "Multiple body biomes (lungs, heart, brain); diverse enemy types; boss fights; upgrade tree. 'Underrated gem' - Reddit",
        "replayability": "Moderate; New Game+ option; 'Great first run, unlikely to replay' - Steam",
    },
    {
        "genre": "2D Platformer", "game": "Desvelado", "tier": "MID", "year": 2024,
        "revenue": "~$189K est.", "review": "360",
        "team_size": "Small indie team (Latin America)",
        "trailer": "Cute spooky atmosphere; light-extinguishing mechanic shown charmingly; gentle haunted lullaby music sets tone immediately",
        "ui": "Minimal and clean; light indicators; nothing distracts from platforming focus",
        "art_style": "Hand-drawn 2D dark pastel; sleepy dreamy aesthetic; cute chibi vampire design; soft shadow palette",
        "feature": "- Precision platformer: extinguish every light in haunted castle\n- Flame-based limited dash (one per flame collected)\n- Ghost enemies re-light lamps (unique tension loop)\n- Wall-jump and standard platforming vocabulary\n- No violence - pure puzzle-platforming",
        "scope": "2-3hr; focused precision platformer; polished single-playthrough; $6.99 price point",
        "content": "Castle environments with varied puzzle rooms; ghost patrol enemies; light-based challenge variety. 'Perfect Halloween game' - Steam",
        "replayability": "Low; experience-focused; 'Finished in one sitting, worth it' - Steam",
    },
    {
        "genre": "2D Platformer", "game": "Zefyr: A Thief's Melody", "tier": "MID", "year": 2025,
        "revenue": "~$252K est.", "review": "337",
        "team_size": "Solo (Mathias Fontmarty, 11 years dev)",
        "trailer": "Beautiful cel-shaded world; Zelda BOTW aesthetic immediately apparent; gentle ocean music; freedom and exploration communicated clearly",
        "ui": "Clean minimal; freedom-focused; no combat HUD clutter; traversal feels unobstructed",
        "art_style": "Cel-shaded 3D; soft watercolor-inspired palette; vibrant floating islands; Zelda BOTW aesthetic with anime cel-shading",
        "feature": "- Free-climbing on almost any surface\n- Open 3D island environments\n- Optional stealth sections\n- Environmental puzzles and collectible quests\n- No fall damage; no health bar; combat optional",
        "scope": "6-10hr; multiple themed islands; 11-year solo development; June 2025 launch",
        "content": "Multiple floating island environments; stealth and puzzle blend; contemplative wordless narrative. 'Solo dev masterwork' - Steam",
        "replayability": "Low-moderate; exploration replay; 'I just want to exist in this world' - Steam",
    },
    {
        "genre": "2D Platformer", "game": "Symphonia", "tier": "MID", "year": 2024,
        "revenue": "~$221K est.", "review": "295",
        "team_size": "Team of ~4 (Sunny Peak) + Headup pub.",
        "trailer": "Beautiful musical atmosphere; violinist launch mechanic shown elegantly; orchestral score evocative; visually distinct from other platformers",
        "ui": "Clean minimal floating HUD; momentum cues subtle; nothing distracts from aesthetic experience",
        "art_style": "Hand-drawn/painterly 2D; muted baroque palette; ornate concert hall and cathedral environments; Baroque visual identity",
        "feature": "- Momentum-based platforming: violin launches Philemon through levels\n- Precision platformer with accessibility options (double jump, slow-mo)\n- 4 distinct themed locations\n- Upgradeable movement abilities\n- Wordless musical narrative\n- Orchestral score integrated into level design",
        "scope": "3-4hr; 4 themed locations; $9.99 price; difficulty can extend to 7hr+",
        "content": "4 baroque/orchestral environments; boss encounters; atmospheric wordless narrative. 'Short but unforgettable' - DualShockers",
        "replayability": "Low; experience-focused; 'Perfect first playthrough' - Metacritic",
    },
    {
        "genre": "2D Platformer", "game": "Windswept", "tier": "MID", "year": 2025,
        "revenue": "~$215K est. (barely broke even per dev)", "review": "287",
        "team_size": "Solo (PeekingBoo / WeatherFell)",
        "trailer": "Colorful retro showcase; two-character buddy gameplay shown; nostalgic SNES energy; heartwarming tone",
        "ui": "Clean 2-character HUD; lives and health readable; minimal but functional for retro design",
        "art_style": "Pixel art SNES-era; bright primary colors; DKC/Kirby/Mario World visual touchstones; cheerful cartoon style",
        "feature": "- 2-character buddy platformer (duck + tortoise with different abilities)\n- 40+ stages across themed worlds\n- Secrets and collectibles in every stage\n- Local co-op split-screen support\n- Retro challenge-first level design",
        "scope": "8-12hr; 40+ stages; co-op mode adds replay; secrets dense per level",
        "content": "40+ stages with themed worlds; hidden collectibles; bonus challenges. 'Best retro platformer 2025' - critics; dev barely broke even despite 97% positive",
        "replayability": "Moderate; co-op replay value strong; time trials; 'More fun with a friend' - Steam",
    },
    {
        "genre": "2D Platformer", "game": "Curse of the Sea Rats", "tier": "MID", "year": 2023,
        "revenue": "~$153K est.", "review": "102",
        "team_size": "Team of ~8 (Petoons Studio, Spain)",
        "trailer": "Hand-drawn art shown prominently; adventurous pirate tone; Metroidvania gameplay highlighted; beautiful but hook unclear",
        "ui": "Standard Metroidvania HUD; map + health + weapon; somewhat cluttered; pacing issues noted by reviewers",
        "art_style": "Hand-painted 2D; vibrant rat pirate aesthetic; detailed character animations; colorful ocean and dungeon setting",
        "feature": "- Metroidvania platformer with 4 unique rat characters\n- Hand-drawn frame-by-frame animation\n- NPC side quests\n- Multiple biomes with different mechanics\n- Full interconnected map exploration",
        "scope": "10-12hr Metroidvania; 4 characters; multiple biomes; boss fights; full map",
        "content": "Multiple pirate biomes; 4 rat protagonists; boss fights; NPC storylines. 'Beautiful art sadly paired with mediocre gameplay' - Metacritic",
        "replayability": "Low; gameplay issues limit replay; 55% positive reflects quality vs. investment gap",
    },

    # ---- FAILURE TIER ----
    {
        "genre": "2D Platformer", "game": "Obsolete", "tier": "FAILURE", "year": 2022,
        "revenue": "<$10K est.", "review": "~30",
        "team_size": "Solo (leFarfelu, first-time dev)",
        "trailer": "Minimal or absent; near-zero marketing presence; no social media footprint",
        "ui": "Monochrome minimal; blob transformation indicator only; no keybinding support (friction noted)",
        "art_style": "Monochrome minimalist pixel art; black-white only palette; isolation atmosphere as visual intent",
        "feature": "- Simple movement/jump/blob transformation mechanic\n- Environmental storytelling through minimalism\n- No combat; exploration only\n- Very limited mechanical vocabulary",
        "scope": "1-2hr; very small scale; first-time developer project",
        "content": "Short single-environment experience; atmospheric OST cited as main strength. 'Evokes isolation well but thin on gameplay' - Steam",
        "replayability": "None; no incentive to replay; experience exhausted on first playthrough",
    },
    {
        "genre": "2D Platformer", "game": "Bushiden", "tier": "FAILURE", "year": 2022,
        "revenue": "<$50K est.", "review": "~50",
        "team_size": "Team of 3 (Pixel Arc Studios)",
        "trailer": "Impressive cyberpunk pixel art in Kickstarter trailers (2018); EA launch buzz failed to convert; incomplete state at launch undermined initial interest",
        "ui": "Unpolished EA state; reviews note incomplete UI/UX",
        "art_style": "2D pixel art cyberpunk ninja; dark neon palette; Ninja Gaiden/Strider 16-bit inspiration; visually distinctive but EA rough",
        "feature": "- Ninja action platformer with cybernetic upgrade system\n- Kickstarter 2018 with ambitious scope\n- EA launch without sufficient content to retain players\n- Metroidvania exploration elements in design",
        "scope": "EA with partial content ~3-4hr; development stretched from 2018 Kickstarter through 2022+ without full release",
        "content": "Incomplete biomes; cyberpunk city setting promising but undelivered. 'Great potential stuck in EA limbo' - Steam reviews",
        "replayability": "None; EA state and low content discourage return; no active community",
    },
    {
        "genre": "2D Platformer", "game": "Leap Year", "tier": "FAILURE", "year": 2024,
        "revenue": "<$15K est.", "review": "~20",
        "team_size": "Solo (unknown)",
        "trailer": "Minimal; concept-driven pitch (collect Feb 2024 calendar pages) without marketing execution",
        "ui": "Unknown; game has near-zero press coverage",
        "art_style": "Pixel art; calendar/time-travel theme; minimal visual identity beyond concept",
        "feature": "- 2D platformer collecting calendar pages (Feb 2024 theme)\n- Concept-first design without mechanical depth\n- Very limited scope",
        "scope": "1-2hr estimated; niche concept without broad appeal",
        "content": "Calendar-themed levels; platformer elements; thematic novelty without content depth",
        "replayability": "None; concept exhausted in first playthrough; no community formed",
    },
    {
        "genre": "2D Platformer", "game": "Generic Pixel Platformer (archetype)", "tier": "FAILURE", "year": 2023,
        "revenue": "<$5K est.", "review": "~10",
        "team_size": "Solo (first-time dev)",
        "trailer": "No distinguishing mechanic in capsule or trailer; looks like one of ~1,960 annual releases; zero social media traction",
        "ui": "Standard 2D platformer HUD; nothing distinguishing the experience",
        "art_style": "Generic pixel art; default Unity/GMS2 asset-adjacent look; no distinctive visual identity",
        "feature": "- Standard jump, run, collect loop\n- No distinguishing mechanic or hook\n- Generic enemy and level design\n- Represents 99%+ of ~1,960 annual 2D platformer releases",
        "scope": "2-4hr; no post-launch support; developer moved on after poor launch",
        "content": "Generic levels with no memorable moments; no story investment; invisible on Steam from launch day",
        "replayability": "Zero; no wishlist momentum; no community; Steam algorithm never surfaces it",
    },
]

sheet_rows = []
tiers      = []
for g in game_rows:
    sheet_rows.append([
        g["genre"], g["game"], g["year"], g["revenue"], g["review"],
        g["team_size"], g["trailer"], g["ui"], g["art_style"],
        g["feature"], g["scope"], g["content"], g["replayability"],
    ])
    tiers.append(g["tier"])

ws.append_rows(sheet_rows, value_input_option="USER_ENTERED")

sh.batch_update({"requests": [{
    "autoResizeDimensions": {
        "dimensions": {
            "sheetId": ws.id,
            "dimension": "ROWS",
            "startIndex": 1,
        }
    }
}]})

all_vals  = ws.get_all_values()
total_rows = len(all_vals)
start_row  = total_rows - len(sheet_rows) + 1

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
    if len(format_requests) >= 25:
        sh.batch_update({"requests": format_requests})
        format_requests = []
        time.sleep(1.2)

if format_requests:
    sh.batch_update({"requests": format_requests})

print(f"Written {len(sheet_rows)} games to tab '{tab_name}'.")

# Export CSV + git commit
BASE     = Path(r"C:\Organized Files\My Game Asset\Game-Research")
data_dir = BASE / "data" / "competitors"
data_dir.mkdir(parents=True, exist_ok=True)
csv_path = data_dir / "2d-platformer.csv"
with open(csv_path, "w", newline="", encoding="utf-8") as f:
    csv.writer(f).writerows(ws.get_all_values())
print(f"Exported: {csv_path.relative_to(BASE)}")

def _git(*args):
    r = subprocess.run(["git"] + list(args), cwd=str(BASE), capture_output=True, text=True)
    if r.stdout: print(r.stdout.strip())
    if r.stderr: print(r.stderr.strip())

today_str = date.today().isoformat()
_git("add", "data/competitors/2d-platformer.csv", "research-log.md")
_git("commit", "-m", f"research: 2D Platformer competitor data - {today_str}")
_git("push", "origin", "main")
print("Git: committed and pushed.")
