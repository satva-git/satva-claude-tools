#!/usr/bin/env bash
# Installs the Satva format skill for Cursor (all projects): ~/.cursor/skills/satva-brand-format
# Usage: curl -fsSL https://raw.githubusercontent.com/satva-git/satva-claude-tools/main/install-cursor.sh | bash
set -euo pipefail
BASE="https://raw.githubusercontent.com/satva-git/satva-claude-tools/main/plugins/satva-brand-format/skills/satva-brand-format"
DEST="${HOME}/.cursor/skills/satva-brand-format"
mkdir -p "$DEST"
for f in SKILL.md satva_brand.py; do
  curl -fsSL "$BASE/$f" -o "$DEST/$f"
done
echo "Satva format skill installed in $DEST"
echo "Restart Cursor (or open a new Agent chat) to load it. Run this command again to update."
