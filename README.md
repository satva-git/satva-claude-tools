# Satva Claude tools

Claude plugin marketplace for Satva Solutions.

## satva-brand-format
Gives every document, PDF, slide deck, Google Doc, Google Slides file and artifact (HTML page) the Satva house format: theme, header, footer and styles.

### Install (one time, works in every project)
```bash
claude plugin marketplace add satva-git/satva-claude-tools
claude plugin install satva-brand-format@satva-tools
```
You can also run the same commands inside Claude Code with `/plugin marketplace add satva-git/satva-claude-tools` and `/plugin install satva-brand-format@satva-tools`.

After that, the format applies automatically in every project. Ask Claude for a BRD, proposal, deck, PDF, Google Doc, Google Slides or web page and it uses the Satva theme, header, footer and styles. There's no command to run: the plugin tells Claude at the start of each session to load the format.

To check it's installed, run `claude plugin list`. `satva-brand-format@satva-tools` should show as enabled. Restart any session that was already open when you installed it.

### Get updates
```bash
claude plugin marketplace update satva-tools
```

### Claude app (claude.ai / desktop)
Upload `satva-brand-format.zip` under Customize > Skills and switch it on, or ask the owner to share the skill with you in the app. The app picks the skill automatically from its description, so no command is needed.

## Updating the format (maintainers)
1. Edit `plugins/satva-brand-format/skills/satva-brand-format/SKILL.md`.
2. Increase `version` in `plugins/satva-brand-format/.claude-plugin/plugin.json`.
3. Run `claude plugin validate .`, then commit and push.
