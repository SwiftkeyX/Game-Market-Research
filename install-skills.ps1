# Install indie game market research skills into Cowork
# Run once in PowerShell, then restart the Claude desktop app.

$src  = "C:\Organized Files\My Game Asset\Game-Research\.claude\skills"
$dest = "C:\Users\ad\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\e9b8689c-9b20-42ce-98cf-b3b63a6dccdb\5ee34502-ed3e-4ed0-8a67-0487274e591b\skills"

$skills = @(
    "indie-game-market-research",
    "competitor-lookup",
    "genre-viability-check",
    "genre-viability-data",
    "revenue-target"
)

foreach ($skill in $skills) {
    $from = Join-Path $src $skill
    $to   = Join-Path $dest $skill
    if (Test-Path $from) {
        Copy-Item -Path $from -Destination $to -Recurse -Force
        Write-Host "Installed: $skill"
    } else {
        Write-Host "Not found: $from"
    }
}

Write-Host ""
Write-Host "Done. Restart the Claude desktop app to load the new skills."
