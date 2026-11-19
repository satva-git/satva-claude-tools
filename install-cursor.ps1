# Installs the Satva format skill for Cursor (all projects): %USERPROFILE%\.cursor\skills\satva-brand-format
# Usage (PowerShell): irm https://raw.githubusercontent.com/satva-git/satva-claude-tools/main/install-cursor.ps1 | iex
$ErrorActionPreference = "Stop"
$base = "https://raw.githubusercontent.com/satva-git/satva-claude-tools/main/plugins/satva-brand-format/skills/satva-brand-format"
$dest = Join-Path $HOME ".cursor\skills\satva-brand-format"
New-Item -ItemType Directory -Force -Path $dest | Out-Null
foreach ($f in @("SKILL.md", "satva_brand.py")) {
    Invoke-WebRequest -UseBasicParsing -Uri "$base/$f" -OutFile (Join-Path $dest $f)
}
Write-Host "Satva format skill installed in $dest"
Write-Host "Restart Cursor (or open a new Agent chat) to load it. Run this command again to update."
