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

### Get updates (Claude Code)
When the format changes, everyone runs:
```bash
claude plugin marketplace update satva-tools
claude plugin update satva-brand-format@satva-tools
```
Then restart Claude Code (close the session and run `claude` again). To check which version you have, run `claude plugin list`.

### Claude app (claude.ai / desktop)
**Install:** go to Customize > Skills, upload `satva-brand-format.zip` and switch it on. Or ask the owner to share the skill with you in the app. The app picks the skill automatically from its description, so no command is needed.

**Update:**
- **If the skill was shared with you in the app:** nothing to do. You get the owner's changes automatically.
- **If you uploaded the zip yourself:** download the latest `satva-brand-format.zip` from this repo. In Customize > Skills, delete your old `satva-brand-format`, then upload the new zip and switch it on.

## Updating the format (maintainers)
1. Edit `plugins/satva-brand-format/skills/satva-brand-format/SKILL.md`. Colours, fonts and the phone and email are in the `brand tokens` block of the helper module near the end of the file.
2. Increase `version` in `plugins/satva-brand-format/.claude-plugin/plugin.json` (for example `1.1.0` to `1.2.0`). Claude Code only offers the update when this number changes.
3. Rebuild the app zip from the repo root:
   ```bash
   rm -f satva-brand-format.zip
   (cd plugins/satva-brand-format/skills && zip -r ../../../satva-brand-format.zip satva-brand-format)
   ```
4. Check it: `claude plugin validate .` should end with `Validation passed`.
5. Commit and push to `main`.
6. Update the owner's copy in the Claude app (Customize > Skills: replace it with the new zip, or ask Claude to update the skill). Everyone it's shared with gets the change automatically.
7. Tell the team to run the two update commands under **Get updates**.
