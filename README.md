# Satva Claude tools

Claude plugin marketplace for Satva Solutions.

## satva-brand-format
Gives every document, PDF, slide deck, Google Doc and Google Slides file the Satva house format: theme, header, footer and styles.

### Install (one time, works in every project)
```bash
claude plugin marketplace add <owner>/<repo>
claude plugin install satva-brand-format@satva-tools
```
You can also run the same commands inside Claude Code with `/plugin marketplace add <owner>/<repo>` and `/plugin install satva-brand-format@satva-tools`.

After that, ask Claude for a BRD, proposal, deck or PDF and it uses the house format automatically.

### Get updates
```bash
claude plugin marketplace update satva-tools
```

### Claude app (claude.ai / desktop)
Upload `satva-brand-format.zip` under Customize > Skills, or ask the owner to share the skill with you in the app.

## Updating the format (maintainers)
1. Edit `plugins/satva-brand-format/skills/satva-brand-format/SKILL.md`.
2. Increase `version` in `plugins/satva-brand-format/.claude-plugin/plugin.json`.
3. Run `claude plugin validate .`, then commit and push.
