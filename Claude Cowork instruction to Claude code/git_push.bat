@echo off
:: git_push.bat
:: ============
:: Commits and pushes the latest research data to GitHub.
:: Run this from the Game-Research folder, or double-click it.
::
:: If prompted for credentials, use your GitHub username +
:: a Personal Access Token (not your password).
:: Create one at: https://github.com/settings/tokens  (scope: repo)

cd /d "%~dp0"

echo.
echo == Git status ==
git status --short

echo.
echo == Staging research files ==
git add data\competitors\roguelike-deckbuilder.csv
git add research-log.md

echo.
echo == Committing ==
git commit -m "research: Roguelike Deckbuilder competitor data - %date:~10,4%-%date:~4,2%-%date:~7,2%"

echo.
echo == Pushing to origin/main ==
git push origin main

echo.
echo Done. Press any key to close.
pause
