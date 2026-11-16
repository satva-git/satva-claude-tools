#!/usr/bin/env bash
# Adds a standing instruction to every Claude Code session so the Satva format applies without a command.
cat <<'JSON'
{"hookSpecificOutput":{"hookEventName":"SessionStart","additionalContext":"Satva house format is ON for this session. Before creating or editing any document, report, BRD, FRD, SRS, SOW, proposal, PDF, Word file, slide deck, PowerPoint, spreadsheet, Excel file, Google Doc, Google Slides, HTML page/artifact or email draft, load the satva-brand-format skill (Skill tool: satva-brand-format:satva-brand-format) and follow it. Do this automatically; the user does not need to ask for the format or run a command."}}
JSON
