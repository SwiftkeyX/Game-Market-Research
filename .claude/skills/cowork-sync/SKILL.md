---
name: cowork-sync
description: "Run pending Cowork scripts — uploads research data to Google Sheets, then commits and pushes to GitHub. Invoke after a scheduled Cowork research run completes."
---

# Cowork Sync

Finishes what Claude Cowork started: uploads new competitor data to Google Sheets, then commits and pushes the research files to GitHub.

**Why this exists:** Cowork runs in a sandboxed environment that blocks outbound Google OAuth calls, so it cannot write to Sheets directly. It saves results as `.xlsx` and `.csv` files and leaves upload scripts here. Claude Code runs on the user's machine where full network access is available, so it handles the Sheets write and git push.

---

## Instructions

### Step 1 — Discover pending sheets scripts

Scan for any `.py` files in:
```
C:\Organized Files\My Game Asset\Game Market\Game-Research\Claude Cowork instruction to Claude code\scripts\sheets\
```

List what you find. If the directory is empty, tell the user there are no pending uploads and stop.

### Step 2 — Run each sheets script

For each `.py` file found, run it with:
```
C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe "Claude Cowork instruction to Claude code\scripts\sheets\<filename>.py"
```

Run from the project root:
```
C:\Organized Files\My Game Asset\Game Market\Game-Research
```

Report success or failure for each script. If a script fails, show the error and ask the user whether to continue with the remaining scripts.

### Step 3 — Delete each sheets script after it runs

These are one-time-use scripts. After each one runs successfully, delete it:
```
Remove-Item "C:\Organized Files\My Game Asset\Game Market\Game-Research\Claude Cowork instruction to Claude code\scripts\sheets\<filename>.py"
```

### Step 4 — Run git_push.bat

After all sheets scripts have run, commit and push the new research data to GitHub:
```
"C:\Organized Files\My Game Asset\Game Market\Game-Research\git_push.bat"
```

Or from the project root terminal:
```
git_push.bat
```

Report the output. If git asks for credentials, tell the user to enter their GitHub username and a Personal Access Token (not their password). They can create one at https://github.com/settings/tokens (scope: `repo`).

---

## Python interpreter

Always use:
```
C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe
```

If a script fails with a missing module error, install dependencies once:
```
C:\Users\ad\AppData\Local\Programs\Python\Python312\python.exe -m pip install gspread google-auth openpyxl
```

## Credentials

Service account key (already in the project directory):
```
C:\Organized Files\My Game Asset\Game Market\Game-Research\genre-viability-data-417b9f28c38e.json
```

Google Sheet:
```
https://docs.google.com/spreadsheets/d/1xAF6wWvhe0E4kBQV0i_DqTu1hvqdy8HL07YZyTtruCw/edit
```

---

## Summary output

After completing all steps, report:
- Which sheets scripts ran and which tabs they wrote to
- Which scripts were deleted
- Whether the git push succeeded and the commit hash
