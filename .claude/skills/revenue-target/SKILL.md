---
name: revenue-target
description: >
  Usage: [team size][genre] — Pick a realistic revenue target and north star game based on team
  size. Optionally anchors to a specific genre for more relevant examples.
  Examples: [Solo][roguelike deckbuilder]  |  [Team of 4]
---

# Revenue Target

Helps the user pick a realistic financial target before committing to production.
The key rule: find games that look *worse* than a typical entry in the genre and confirm
they still hit the target. If strong-looking competitors underperform, that is a genre red flag.

## Arguments

```
/revenue-target [team size][genre]   # team size required; genre optional
/revenue-target [Solo]               # derive target for solo dev, no genre filter
/revenue-target [Team of 4][auto-battler]  # team + genre for genre-specific examples
```

Valid team sizes: Solo, Duo, Team of 3–4, Team of 5+

---

## Revenue Bracket Table

| Team | Realistic Target | North Star Example |
|---|---|---|
| Solo | $100K – $500K | Slice & Dice (solo dev, abstract art, $1M+) |
| Duo | $200K – $800K | Into the Breach (2 people, $10M+) |
| Team of 3–4 | $500K – $2M | 9 Kings (small team, $2.2M launch month) |
| Team of 5+ | $1M – $5M | Monster Train (~10 people, $5M+) |

---

## Workflow

### Step 1 — Derive target from team size

Match the user's team size to the bracket table above. Output the realistic target range.

**North Star note**: North Star examples are outliers chosen to show what's *possible* —
they are NOT the expected outcome. The realistic target is the number to plan around.

### Step 2 — Find genre-specific examples (if genre was provided)

Run a web search to find 2–3 games in the genre that:
- Were made by a team of similar size
- Hit a revenue that falls **within** (not above) the realistic target

```
web_search: [genre] steam indie [team size] developer revenue $[low end of target]
web_search: [genre] steam [team size] solo dev success story
```

For each example found, collect:
- Game name, estimated revenue, team size, year
- Brief observation: what did they do right? What would a worse-looking game still have achieved?

### Step 3 — Red flag check

Ask: do the genre's HIGH-tier competitors look significantly better than the realistic target
would imply? If top competitors have 7+ scores on Art, Game Feel, and Content but the genre's
median revenue is still only $200K — that is a warning sign that the genre punishes anything
below the quality bar severely.

Flag this if found:
> ⚠️ The quality bar in this genre is high relative to the revenue upside. Strong competitors
> still land in mid-tier revenue. Scope carefully.

### Step 4 — Output

```
## Revenue Target: [Genre] for [Team Size]

**Recommended Target**: $X – $Y
**North Star Game**: [Game] ([team size], [revenue estimate])
  ↳ Note: This is an outlier. Plan around the bracket, not the north star.

**Genre-specific benchmarks** (games your team size could realistically match):
| Game | Revenue | Team | Year | Why it's a useful benchmark |
...

**Risk signal**: [None / Flag if quality bar is disproportionate to revenue]
```
