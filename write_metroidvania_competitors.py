import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import gspread
from google.oauth2.service_account import Credentials
from datetime import date
import time

CREDENTIALS_FILE = r"C:\Organized Files\My Game Asset\Game-Research\genre-viability-data-417b9f28c38e.json"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1p7vPIIR8imPAZUZekbIuDW7Qfg0Jfc9xa6eEeaSv5Us/edit?usp=sharing"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
client = gspread.Client(auth=creds)
sh     = client.open_by_url(SHEET_URL)

COLOR_HEADER  = {"red": 0.812, "green": 0.886, "blue": 0.953}
COLOR_HIGH    = {"red": 0.851, "green": 0.918, "blue": 0.827}
COLOR_MID     = {"red": 1.0,   "green": 0.949, "blue": 0.800}
COLOR_FAILURE = {"red": 1.0,   "green": 0.902, "blue": 0.902}
COLOR_WHITE   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

HEADERS = ["Genre","Game","Category","Year","Subgenre","Est. Revenue","Reviews","Score","Price","Team","Art Style","Notes"]

# Get or create Competitors tab
ws_list = [w for w in sh.worksheets() if "Competitors" in w.title]
if ws_list:
    ws = ws_list[0]
else:
    ws = sh.add_worksheet(title="Competitors", rows=500, cols=15)
    ws.append_row(HEADERS)
    ws.format("A1:L1", {"backgroundColor": COLOR_HEADER, "textFormat": {"bold": True}})

rows = [
    # ===== HIGH =====
    ["Metroidvania","Hollow Knight","HIGH","2017","Pure Exploration Metroidvania","~$30M+ est.","534000","Overwhelmingly Positive 97%","$14.99","Team Cherry (Duo dev)","Hand-drawn 2D","THE genre benchmark. Every new metroidvania is compared to this. 2 main devs built a 40-hour masterpiece. Still selling daily in 2025. Sets the quality ceiling: no shortcuts on world depth, art, or feel."],
    ["Metroidvania","Dead Cells","HIGH","2018","Roguelike Action Metroidvania","~$50M+ est.","178000","Overwhelmingly Positive 97%","$24.99","Motion Twin (~12)","Stylized pixel art","Defined the roguelike-metroidvania subgenre. Long EA period built community before 1.0. Still active with DLC in 2025. Proof: fluid action feel + procedural structure = lasting player base."],
    ["Metroidvania","Animal Well","HIGH","2024","Puzzle-Exploration Metroidvania","~$21.3M est.","23254","Overwhelmingly Positive 96%","$24.99","Billy Basso (Solo dev)","Painterly glowing pixel","Solo dev 7-year debut. No combat — pure exploration and puzzle box. 1.2M copies sold. Published by Bigmode. Proof: unique execution at high quality beats genre convention. Best solo dev HIGH benchmark."],
    ["Metroidvania","Ender Magnolia: Bloom in the Mist","HIGH","2025","Dark Fantasy Combat Metroidvania","~$17.3M est.","17600","Overwhelmingly Positive 97%","$24.99","Adglobe / Live Wire (small)","Hand-drawn anime dark fantasy","Sequel to Ender Lilies. One of highest Metacritic scores of 2025. Hand-drawn anime + precise parry combat is a proven formula. Quality execution of a known structure wins in this genre."],
    ["Metroidvania","Nine Sols","HIGH","2024","Parry-Action Dao-Punk Metroidvania","~$12M est.","18000","Overwhelmingly Positive 96%","$29.99","Red Candle Games (~10)","Hand-drawn Asian cel-shaded","Unique Taoist/sci-fi setting + deep lore + precise parry combat. 800K+ copies across all platforms. Proof: strong cultural identity differentiates in a crowded genre more than gameplay innovation alone."],
    ["Metroidvania","Blasphemous 2","HIGH","2023","Dark Religious Action Metroidvania","~$8.7M est.","18289","Very Positive 90%","$29.99","The Game Kitchen (~25)","Gothic dark pixel / hand-drawn hybrid","Sequel ceiling: dropped from Overwhelmingly Positive to Very Positive vs the original. Triple-weapon system is the design innovation. Lesson: sequels split communities; execution must clearly improve or score suffers."],
    ["Metroidvania","Pseudoregalia","HIGH","2023","3D Movement-Focused Metroidvania","~$2.3M est.","13499","Overwhelmingly Positive 97%","$5.99","rittzler (Solo dev)","N64-era low-poly 3D","Game jam origin, finished in a few months. 496K copies at $6. Community darling. Proof: movement feel beats visual fidelity when the physics feel exceptional. Cheapest-price solo HIGH benchmark in the genre."],
    # ===== MID =====
    ["Metroidvania","Gestalt: Steam & Cinder","MID","2024","Steampunk Narrative Action Metroidvania","~$2.64M Boxleiter est.","1758","Mostly Positive 79%","$19.99","Metamorphosis Games (small)","Hand-drawn steampunk 2D","Beautiful art and praised narrative but 79% Mostly Positive ceiling. Now 85% off on discount sites. Art alone does not drive sales when the core gameplay loop is generic. Lesson: visual style must be paired with an equally strong mechanical hook."],
    ["Metroidvania","Moonscars","MID","2022","Soulslike Dark Metroidvania","~$1.76M Boxleiter est.","1170","Very Positive 82%","$19.99","Black Mermaid (small)","Dark hand-drawn gothic 2D","Quality soulslike-metroidvania hybrid. 82% VPos is the MID ceiling when competing against Dead Cells and Hollow Knight. Soulslike-metroidvania is increasingly crowded — differentiation beyond difficulty level is required to break out."],
    ["Metroidvania","Bo: Path of the Teal Lotus","MID","2024","Japanese Folklore Platforming Metroidvania","~$1.95M Boxleiter est.","1301","Mostly Positive 82%","$19.99","Squid Shock Studios (debut, ~5)","Hand-drawn Japanese watercolor 2D","Debut studio, praised for art. Criticized for pacing and combat feel gaps. 82% MPos with beautiful watercolor art shows that visual quality without tight execution caps out in MID. Debut team lesson."],
    ["Metroidvania","Souldiers","MID","2022","Retro Tactics-Platformer Metroidvania","~$1.3M est.","802","Mostly Positive 74%","$19.99","Retro Forge (~10)","Retro pixel RPG art","Nearly killed by a brutally difficult and buggy launch (near Mostly Negative). Patches rescued it to 74%. 74% is the permanent ceiling: first impressions define the Steam score forever. Lesson: launch state cannot be undone by patches."],
    ["Metroidvania","HAAK","MID","2022","Cyberpunk Action Metroidvania","~$434K Boxleiter est.","386","Very Positive 92%","$14.99","Bura Games (China indie)","Cyberpunk pixel art","92% VPos — one of the best scores in MID. Only 386 reviews = near-zero marketing investment. Cancelled planned Xbox port due to Microsoft platform bugs. Proof: great reception without visibility permanently caps revenue at MID."],
    ["Metroidvania","Elderand","MID","2023","Castlevania-Inspired Classic Metroidvania","~$595K Boxleiter est.","529","Mostly Positive 76%","$14.99","Mantra Games (~3-4)","Castlevania-style dark pixel","Clear Castlevania inspiration without clearly improving on it. Genre-literate players recognize clone territory. 76% shows the audience knows the reference. Lesson: your inspiration game must be clearly surpassed, not just matched."],
    ["Metroidvania","Lone Fungus","MID","2023","Mushroom World Classic Metroidvania","~$1.09M Boxleiter est.","728","Very Positive 85%","$19.99","Bjorn Jacobsen (Solo dev)","Cute stylized pixel","Solo dev with dedicated fan base who shipped a sequel (Melody of Spores). 85% VPos is respectable but 728 reviews shows the ceiling when marketing is minimal and the setting is niche. Proof that solo dev sustainability is possible without a breakout."],
    # ===== FAILURE =====
    ["Metroidvania","The Last Case of Benedict Fox","FAILURE","2023","Lovecraftian Puzzle Metroidvania","~$1.7M Boxleiter est. (dead playerbase)","755","Mostly Positive 71%","$29.99","Plot Twist Games (small Polish)","Dark 2.5D hand-drawn","Dead playerbase: 31 concurrent players within months of launch. Game Pass release brought traffic but design was too obtuse to retain players. 71% MPos = score of confusion, not hate. Lesson: Game Pass launch masks commercial failure and does not build a community."],
    ["Metroidvania","Curse of the Sea Rats","FAILURE","2023","Pirate Ratoidvania","~$300K Boxleiter est.","~200","Mixed 57%","$19.99","Petoons Studio (small Spanish)","Hand-drawn pirate rat 2D","Beautiful hand-drawn animation but repetitive combat and poor metroidvania design drove Mixed 57%. Art carried it; the gameplay did not. Mixed = death in a genre where Very Positive is the floor for visibility. Lesson: art cannot compensate for weak gameplay systems."],
    ["Metroidvania","Tales of Kenzera: ZAU","FAILURE","2024","African Mythology Action Metroidvania","~$320K est. (commercial failure)","~915","Very Positive 83%","$19.99","Surgent Studios (~15, EA Originals)","Hand-painted African mythology 2D","83% VPos but only 287 peak concurrent players and ~$320K total revenue. EA-published and free on PS Plus day one — diluted Steam installs. Led to studio restructuring. Lesson: PS Plus launch cannibalizes PC sales; review score does not equal commercial success."],
    ["Metroidvania","INAYAH - Life after Gods","FAILURE","2025","Post-Apocalyptic Action Metroidvania","~$307K Boxleiter est.","~205","Mostly Positive 73%","$19.99","Small indie team","Hand-drawn post-apocalyptic 2D","73% MPos from ~205 reviews = invisible on Steam. Distinctive art and unique setting, but launched without building a wishlist audience. Lesson: no wishlist strategy = no launch momentum regardless of quality. The algorithm needs a minimum launch signal."],
    ["Metroidvania","Vigil: The Longest Night","FAILURE","2020","Dark Horror Metroidvania","~$1.78M Boxleiter est. (launch bugs)","~1400","Mostly Positive 77%","$16.99","Glass Heart Games (China, small)","Dark hand-drawn gothic 2D","Launched with severe performance issues and bugs that drove early reviews negative. Patched over time but first-impression score never recovered. 77% MPos is the permanent ceiling of a bad launch. Lesson: a buggy launch permanently caps your score and kills your launch window."],
    ["Metroidvania","Aeterna Noctis","FAILURE","2022","Dark Epic Action Metroidvania","~$2.97M Boxleiter est.","~1584","Mostly Positive 76%","$24.99","Aeternum Game Studios (~15, Spanish)","Dark hand-drawn epic 2D","4-year ambitious project attempting to surpass Hollow Knight. 76% MPos despite strong visuals shows the Hollow Knight quality bar is exactly that — Hollow Knight level, not close-to-it. Now 80% off on discount sites. Lesson: Hollow Knight ambition requires Hollow Knight execution."],
    ["Metroidvania","Pronty","FAILURE","2023","Underwater Sci-Fi Metroidvania","~$150K Boxleiter est.","~100","Very Positive ~80%","$19.99","Boons and Curses (tiny, Belgian)","Stylized underwater sci-fi 2D","Quality underwater sci-fi concept with decent reception but launched with zero visibility — no wishlist campaign, no community, no marketing presence. Even 80% positive from ~100 reviews = invisible on Steam. Lesson: the algorithm requires a minimum wishlist threshold at launch; without it, quality alone cannot surface the game."],
]

ws.append_rows(rows, value_input_option="USER_ENTERED")
print(f"Appended {len(rows)} rows.")

# Color by category
all_rows = ws.get_all_values()
total = len(all_rows)
start = total - len(rows) + 1

format_requests = []
for i, row in enumerate(rows, start=start):
    cat = row[2]
    color = {"HIGH": COLOR_HIGH, "MID": COLOR_MID, "FAILURE": COLOR_FAILURE}.get(cat, COLOR_WHITE)
    format_requests.append({
        "repeatCell": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": i - 1,
                "endRowIndex": i,
                "startColumnIndex": 0,
                "endColumnIndex": 12,
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

print("Colors applied.")
print(f"Written {len(rows)} Metroidvania competitors to sheet.")
