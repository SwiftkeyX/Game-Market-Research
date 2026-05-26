# CLAUDE.md

This file provides guidance to Claude when working with files in this repository.

## What This Directory Is

A self-contained indie game market research toolkit. The central skill helps you understand
the gaming market as an indie developer — what genres are healthy, who's winning in each space,
and what the competitive landscape actually looks like.

## Main Skill

The primary skill lives at `.claude/skills/indie-game-market-research/SKILL.md`.
Invoke it with `/indie-game-market-research`.

### Two modes

**Mode 1 — Market Overview** (no arguments):
```
/indie-game-market-research
```
Ranks all tracked genres by viability (GO / CAUTION / AVOID), hit rate, trend, and revenue range.
Use this to orient yourself across the whole indie market.

**Mode 2 — Genre Deep Dive** (genre provided):
```
/indie-game-market-research [roguelike deckbuilder]
```
Researches 21 real games in that genre (7 HIGH / 7 MID / 7 FAILURE) and collects:
Genre | Game | Year | Revenue | Reviews | Team Size | Art Style | Game Feel | Features | Scope | Content | Replayability

Saves everything to a genre-specific Google Sheets tab and an Excel file.

## Sub-skills

| Skill | Purpose |
|---|---|
| `genre-viability-data` | Manages the genre ratings sheet (view / update / add / remove genres) |
| `genre-viability-check` | Quick check on a single genre's viability verdict, hit rate, and trend |
| `competitor-lookup` | Full genre deep dive — finds 21 games and collects all descriptive data |
| `revenue-target` | Standalone: maps team size to a realistic revenue bracket (optional, not in main flow) |
| `gameplay-review` | Deprecated — merged into competitor-lookup |

## Google Sheets

All data writes to:
https://docs.google.com/spreadsheets/d/1p7vPIIR8imPAZUZekbIuDW7Qfg0Jfc9xa6eEeaSv5Us/edit

Credentials: `genre-viability-data-417b9f28c38e.json` (in this directory)

## Reference Files

- `references/subgenre-guide.md` — Per-subgenre breakdown with examples
- `references/competitor-examples.md` — Pre-researched data for common genres
- `references/quality-benchmark-guide.md` — Evidence sources for each data dimension
