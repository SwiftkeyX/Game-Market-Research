# Google Sheet — Column Guide

Plain-language explanation of every column in every tab.
What it means, where it comes from, how it is calculated.

Sheet: https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw

---

## Tab: Genre Viability Ratings (GO / CAUTION / AVOID) (and v1 / v2 snapshots)

This is the main decision table. Each row is one game genre.
You read it to decide whether a genre is worth building a game in right now.

The **v1 / v2** tabs are frozen copies of this same table saved at a point in time.
They exist so you can see how ratings changed between updates. Never edit them.

---

### Verdict

**What it is:** A three-level rating of whether to make a game in this genre.

| Value | Meaning |
|---|---|
| GO | The genre has a healthy hit rate right now. Entering is reasonable. |
| CAUTION | The genre has problems — too crowded, declining, too expensive to make, or too risky solo. Proceed carefully or skip. |
| AVOID | The genre almost never produces profitable indie games. Do not enter. |

**Where it comes from:** Chris Zukowski's annual hit-rate analysis at howtomarketagame.com, cross-checked with Steam release data. Updated manually when new data is available.

**How it is decided:** Mostly based on Hit Rate (see below) combined with trend direction. A genre can be CAUTION even with a decent hit rate if the trend is steeply declining.

---

### Genre

**What it is:** The category of game being rated (e.g. Roguelike Deckbuilder, Idle / Incremental).

**What to know:** These are Steam tags and market categories, not strict game design definitions. Two games can have the same gameplay but different genre tags, which affects which audience finds them. Genre here means "how Steam players and media label it."

---

### Solo Target

**What it is:** A realistic revenue range for a solo developer (one person making the game alone) who ships a finished product in this genre.

**Example:** `$100K–$500K` means a solo dev making a good game in this genre can realistically expect to earn somewhere in that range over its lifetime on Steam.

**Where it comes from:** Estimated from real solo-dev games in each genre. It is not guaranteed — it is the range where most successful solo games land, excluding outliers.

**What N/A means:** `N/A (solo skip)` means the genre is not viable solo. Either the scope is too large (Colony Sim takes 2–4 years minimum) or the quality bar requires a team.

---

### Team of 4

**What it is:** Same as Solo Target but for a small team of roughly 4 people.

A team can make a bigger, more polished game, so the revenue ceiling is higher. A small team also has higher costs (salaries or time investment), so the minimum needed to break even is also higher.

---

### Hit Rate

**What it is:** The percentage of games released in this genre that reach 1,000+ Steam reviews.

**Why 1,000 reviews matters:** Reaching 1,000 reviews on Steam roughly corresponds to earning enough revenue to sustain an indie studio. It is the industry benchmark for "this game did not fail commercially."

**Example:** `5.1%` means about 5 out of every 100 games released in this genre hit 1,000+ reviews. The other 95 do not.

**Where it comes from:** howtomarketagame.com annual analysis. Chris Zukowski counts all Steam releases per genre per year and tracks how many cross the 1,000-review threshold.

**What the unusual values mean:**
- `Phase 3 emerging` — the genre is new enough that hit rate hasn't stabilized yet; it's growing fast (see Notes)
- `Top Q1 2026` — too new to have a percentage; it had the most hits in a single quarter
- `Consistent hits` — the genre reliably produces winners each year without a clean percentage
- `Gatekept` — the genre is dominated by established publishers or back-catalog; indie hit rate is near zero even if total genre releases look healthy
- `Skewed (Asian pub)` — the published percentage is inflated by Chinese publishers releasing in bulk; the real Western indie hit rate is much lower

---

### Trend

**What it is:** Whether the genre's hit rate is getting better or worse compared to the previous year.

| Value | Meaning |
|---|---|
| Growing / Rising | More hits per year than last year. Good time to enter. |
| Stable / Consistent | Hit rate has not changed much. Neither a red flag nor a green flag. |
| Declining / Cooling | Fewer hits than last year. The window may be closing. |
| Emerging | Brand new category with rapid early growth. Higher risk, higher upside. |
| Saturated | Too many games released; the audience is split too thin. |
| Reviving | Was declining or AVOID, now showing new hits. Watch before committing. |
| Dropped from top 5 | Was a top genre last year; no longer in the leading group. |

---

### Notes

**What it is:** One or two sentences of extra context that the numbers alone do not capture.

Common things you will find here:
- Real game names that hit in the genre recently (proof the genre is alive)
- Warnings about what the genre requires (e.g. "3D physics exponentially harder")
- Advice on how to approach the genre (e.g. "need a hook beyond just farming")
- Why a CAUTION rating exists despite a decent hit rate

---

## Tab: Research Session History

A log of every research session run with this toolkit. One row per session.
It is written automatically when you run `/indie-game-market-research`.

---

### Date

The date the research was run. Format: YYYY-MM-DD.

---

### Genre

The genre that was researched in that session.

---

### Verdict

The GO / CAUTION / AVOID verdict from the Current Ratings tab at the time of the session.

---

### Revenue Target

The revenue range recommended for the team size that was researched.
Example: `$100K–$400K (Solo) / $300K–$1.5M (Team4)`

---

### Hit Rate

The hit rate value from Current Ratings at the time of the session.

---

### Trend

The trend value from Current Ratings at the time of the session.

---

## Tab: All Competitor Games by Genre

Each row is one real game that was researched as a competitor in a genre.
The full table has 21 games: 7 HIGH, 7 MID, 7 FAILURE.

---

### Genre

Which genre research session this game belongs to. Used to group rows when multiple genres have been researched.

---

### Game

The game's name exactly as listed on Steam.

---

### Category

Which tier this game falls into. Tiers are based on review count and commercial outcome — **not revenue**.

| Value | Criteria |
|---|---|
| HIGH | 1,000+ reviews. The game clearly worked commercially. |
| MID | 100–999 reviews. The game shipped and has players, but did not break out. |
| FAILURE | Fewer than 100 reviews, OR the game was abandoned, delisted, or shut down despite review count. |

**Why a FAILURE game can still show high estimated revenue:** The revenue figure is an estimate calculated from reviews × price. A game with 2,000 reviews at $30 will show ~$4.5M estimated revenue even if the game was commercially unsuccessful for its team size and budget (e.g. Hyper Light Breaker cost far more to make than it earned).

---

### Year

The year the game released out of Early Access (or launched if it skipped Early Access).

Used to judge relevance — games released 5+ years ago are included only if they define the genre or are an important cautionary example. Recent games (last 3–4 years) are prioritized.

---

### Subgenre

A more specific description of what kind of game this is within the genre.

**Why it matters:** Two games in the same genre (e.g. Roguelike) can be very different products targeting different audiences. "Action Roguelite" and "Traditional Roguelike RPG" are both Roguelikes but compete in different sub-markets.

---

### Est. Revenue

**What it is:** A rough estimate of total USD revenue the game has earned on Steam.

**Where it comes from (in order of preference):**
1. games-stats.com — aggregates Steam data to estimate revenue per game
2. VGInsights — similar data source, slightly different methodology
3. Boxleiter formula (when no direct data is available):

```
Est. Revenue = review_count × price × 75
```

The formula works like this:
- Reviews on Steam are roughly 1/30th of copies sold (every 30 buyers leaves 1 review, on average)
- So: reviews × 30 = estimated copies sold
- Copies sold × price = gross revenue
- Multiply by 75 instead of 30 because Steam takes a 30% cut, returns vary, and the ratio fluctuates

**Important limitations:**
- This is an estimate, not an official figure. It can be off by 50% or more for individual games.
- It does not account for development costs. High revenue does not mean the game was profitable.
- F2P games, games with lots of DLC, or games with heavy discounting will have inaccurate estimates.
- Notes like `(development ended)` or `(delisted)` explain why a game with high estimated revenue is still a FAILURE.

---

### Reviews

Total number of Steam reviews the game has received.

**Where it comes from:** The Steam store page. Also visible on SteamDB.

**Why it matters more than revenue:** Review count is the most reliable public signal of a game's commercial success. Revenue estimates can be wrong; review counts are directly observable and correlate closely with copies sold.

---

### Score

The Steam review score shown on the store page.

**Format:** `[Label] [Percentage]` — e.g. `Very Positive 92%`

| Steam Label | Percentage Range | What it means |
|---|---|---|
| Overwhelmingly Positive | 95%+ | Players love it. Almost no complaints. |
| Very Positive | 80–94% | Players are happy. Normal amount of criticism. |
| Mostly Positive | 70–79% | More good than bad, but notable issues exist. |
| Mixed | 40–69% | Divided player base. Often signals a core problem. |
| Mostly Negative | below 40% | Players are unhappy. The game has serious issues. |

A Mixed score (40–69%) is usually a warning sign. Most successful indie games land at Very Positive or higher.

---

### Price

The USD launch price of the game on Steam.

**Why launch price matters:** Price is baked into the revenue estimate formula. A $10 game needs 3× more copies sold than a $30 game to earn the same revenue. It also signals what market the developer was targeting.

---

### Team

How many people made the game.

**Where it comes from:** Developer websites, LinkedIn, interviews, Steam developer page, or press coverage. If the team size is unknown, it is listed as `Unknown`.

**Common values:**
- `Solo dev` — one person made the entire game
- `Duo dev` — two people
- `Team of N` — a small team of roughly N people
- `Studio name (~N)` — a named studio with approximately N staff

**Why it matters:** Team size directly affects how relevant a game is as a benchmark. A solo developer should not compare themselves to a 20-person studio game.

---

### Art Style

A 2–4 word description of the game's visual style.

**Examples:** `Pixel art`, `Hand-drawn 2D`, `Low-poly 3D`, `ASCII / retro pixel`, `Stylized 3D cel-shaded`

**Where it comes from:** Store page screenshots.

**Why it matters:** Art style affects both development cost and audience expectations. A pixel art roguelike is compared to other pixel art games. A hand-drawn game raises the visual quality bar significantly.

---

### Notes

One or two sentences explaining why this game succeeded or failed, and what lesson to take from it.

This is the most important column for learning. Do not skip it.

---

## Tab: Quality Scores

Each row is the same game from the Competitors tab, now scored across 8 quality dimensions.
The goal is to understand what quality level you need to compete in this genre.

---

### Genre / Game / Category / Date

Same as the Competitors tab. Date is when the scoring was done.

---

### The 8 scored dimensions

All scores are 1–10. Higher is better. See the full scoring rubric at `references/quality-benchmark-guide.md`.

There are two types of dimensions:

**Observable (scored from the Steam store page — anyone can verify these):**

| Column | What is being scored |
|---|---|
| Art | How good the game looks in screenshots. Does it look like a finished, professional product? |
| Store Page | How well the store page communicates what the game is and why it is fun. Players decide in ~10 seconds. |
| Trailer | How good the launch trailer is. Does it show the fun immediately? Does it make you want to play? |
| UI | How clean and polished the in-game interface looks. Bad UI signals an unfinished game. |
| Hook | How unique and easy to explain the game's core concept is. Can you describe it in one sentence? |

**Review-required (scored by reading player reviews — cannot be assessed from screenshots alone):**

| Column | What is being scored |
|---|---|
| Game Feel | How satisfying it feels to play. Controls, responsiveness, audio/visual feedback on actions. |
| Content | How much content exists. Measured in hours of play and whether players feel they got value for money. |
| Replayability | Whether players keep coming back after their first win. Build variety, run diversity, "one more run" drive. |

---

### Feel Source / Content Source / Replay Source

For Game Feel, Content, and Replayability — these columns hold the actual player quote used to justify the score.

**Format:** `"[quote from player]" — Steam review` or `— Reddit r/[subreddit]` or `— [outlet] review`

**Why this exists:** These three dimensions cannot be reliably guessed from a store page. A score without a citation is an opinion. A score with a real player quote is evidence.

**What N/A means:** Not enough public reviews were found to score this dimension reliably. A missing score is honest; a made-up score would be misleading.

---

### Total

The sum of all 8 dimension scores. Maximum possible is 80 (10 × 8).

**How to use it:** Compare the Total of HIGH games to the Total of MID and FAILURE games in the same genre. The gap between HIGH and MID tells you how much quality matters in this genre. A large gap means quality is heavily rewarded; a small gap means other factors (timing, marketing, community) dominate.

---

### Flag

A label based on Total score and Category tier.

| Flag | Criteria | Meaning |
|---|---|---|
| STRONG | HIGH tier game | This game succeeded and has strong quality scores. This is what you are competing against. |
| DECENT | MID tier, total ≥ 50 | Solid game that didn't break out. Quality is there but something else held it back. |
| WEAK | MID tier, total < 50 | Mediocre quality that landed in the middle tier. Getting by, not succeeding. |
| LOW | FAILURE tier | Low quality game that failed. Study what went wrong. |

**Note:** A FAILURE game can sometimes score DECENT (like Realm of Ink, which scored 58 despite being delisted). This means the game failed for non-quality reasons — business issues, publishing rights, platform problems.

---

## Glossary of Terms Used Across the Sheet

| Term | Plain English |
|---|---|
| Early Access (EA) | A game released on Steam before it is fully finished. Players pay to play an incomplete version while the developer keeps building it. |
| CCU | Concurrent users — how many people are playing the game at the same time. A useful measure of an active playerbase, especially for multiplayer games. |
| F2P | Free to Play — the game costs nothing to download. Revenue comes from optional purchases inside the game instead. |
| Hit rate | The percentage of games in a genre that reach 1,000+ Steam reviews (the commercial success threshold). |
| Outlier | A game that performed far above or below what is normal for its genre or team size. Outliers are mentioned for context but should not be used as planning targets. |
| Boxleiter formula | The math used to estimate a game's revenue when no official data is available: reviews × price × 75. Named after the analyst who popularized it. |
| Q1 / Q2 / Q3 / Q4 | Quarters of the calendar year. Q1 = January–March, Q2 = April–June, Q3 = July–September, Q4 = October–December. |
| Wishlist | A Steam feature where players save a game to buy later. Wishlists before launch predict launch revenue. |
| Solo dev | A single person who made the entire game — art, code, design, and sound. |
| Back-catalog | Older, already-released games from established publishers that compete with new releases in the same genre. |
| Rollback netcode | A technical standard for online multiplayer games that makes the connection feel smooth even with lag. Required in competitive games like fighting games. |
| Postmortem | A developer's own written analysis of why their game succeeded or failed, usually published after release. |
