WHY THIS FOLDER EXISTS
======================

It is expensive for Cowork to write to Google Sheets, so we leave this task to Claude Code instead.

When Claude (Cowork) runs a scheduled research task, it saves results locally as .xlsx and .csv files.
Claude Code then picks up the scripts in this folder to push that data to Google Sheets and GitHub.

FOLDER STRUCTURE
----------------
scripts/
  excel/    - Build .xlsx files from research data
  sheets/   - One-time Google Sheets upload scripts (delete after running)
  utils/    - Reusable tools: read viability row, clean tabs, migrate sheet, refine genres

git_push.bat
          - Commit and push latest research data to GitHub
