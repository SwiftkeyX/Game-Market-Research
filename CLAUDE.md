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

Saves research to a local Excel file and CSV. Google Sheets write is handled separately by Claude Code (see below).

## Sub-skills

| Skill | Purpose |
|---|---|
| `genre-viability-data` | Manages the genre ratings sheet (view / update / add / remove genres) |
| `genre-viability-check` | Quick check on a single genre's viability verdict, hit rate, and trend |
| `competitor-lookup` | Full genre deep dive — finds 21 games and collects all descriptive data |
| `revenue-target` | Standalone: maps team size to a realistic revenue bracket (optional, not in main flow) |
| `gameplay-review` | Deprecated — merged into competitor-lookup |

## Google Sheets

All data targets:
https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit

Credentials: `genre-viability-data-417b9f28c38e.json` (in this directory)

⚠️ Cowork CANNOT write to Google Sheets directly — the sandbox blocks Google OAuth.
Reading the sheet is fine via the Google Drive MCP connector.
Writing is delegated to Claude Code via the handoff folder below.

## Handoff to Claude Code
This is instruction to Claude Cowork only. If you are Claude code, you can skip this.

After every `/indie-game-market-research [genre]` run, Cowork must place two files into:
`Claude Cowork instruction to Claude code/scripts/sheets/`

1. The generated Google Sheets upload script (e.g. `upload_[genre-slug]_to_sheets.py`)
   — mark it ⚠️ ONE-TIME USE — DELETE AFTER RUNNING at the top of the file
2. A reminder that `git_push.bat` (already in the folder) should be run after the sheet upload

Claude Code then runs those scripts on the user's Windows machine where full network
access is available. See `Claude Cowork instruction to Claude code/INSTRUCTIONS.txt`
for the full workflow.

## Reference Files

- `references/subgenre-guide.md` — Per-subgenre breakdown with examples
- `references/competitor-examples.md` — Pre-researched data for common genres
- `references/quality-benchmark-guide.md` — Evidence sources for each data dimension
