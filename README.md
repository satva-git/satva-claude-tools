# Satva Claude tools

Claude plugin marketplace for Satva Solutions.

## satva-brand-format
Gives every document, PDF, slide deck, spreadsheet, Google Doc, Google Slides file and artifact (HTML page) the Satva house format (theme, header, footer and styles) and Satva writing rules, so the output reads as Satva-made, not AI-generated.

### Install (one time, works in every project)
```bash
claude plugin marketplace add satva-git/satva-claude-tools
claude plugin install satva-brand-format@satva-tools
```
You can also run the same commands inside Claude Code with `/plugin marketplace add satva-git/satva-claude-tools` and `/plugin install satva-brand-format@satva-tools`.

After that, the format applies automatically in every project. Ask Claude for a BRD, proposal, deck, PDF, Google Doc, Google Slides or web page and it uses the Satva theme, header, footer and styles. There's no command to run: the plugin tells Claude at the start of each session to load the format.

To check it's installed, run `claude plugin list`. `satva-brand-format@satva-tools` should show as enabled. Restart any session that was already open when you installed it.

### First-time setup on your computer (Claude Code)
The plugin builds the files with Python on your own computer, so it needs a few free tools there:

1. **Python 3** (python.org). Check with `python --version` (on macOS: `python3 --version`).
2. **Python libraries:**
   ```bash
   python -m pip install python-docx python-pptx openpyxl Pillow
   ```
   If pip says "externally-managed-environment", add `--user` to the end.
3. **PDF export:** install LibreOffice (libreoffice.org). On Windows or macOS with Microsoft Word you can instead run `python -m pip install docx2pdf`.
4. **Fonts:** install **Mulish** and **Lexend** from fonts.google.com so the files and PDFs show the right fonts.

To check everything at once, ask Claude Code: "check the Satva format setup". It runs the plugin's checker and installs anything that's missing.

### If the format isn't applied
1. Run `claude plugin list`. `satva-brand-format@satva-tools` must be listed and enabled, at version 1.4.0 or later. If not, run the update commands below.
2. Start a **new** Claude Code session after installing or updating. Sessions that were already open don't pick up plugins.
3. Ask for the output by type, for example "create a BRD for …", "make a deck on …" or "export this as a PDF". The plugin reminds Claude to use the format whenever a request mentions a document, deck, PDF, spreadsheet, page or email.
4. To force it in any session, type `/satva-brand-format:satva-brand-format` before your request.
5. If Claude reports a missing library or PDF converter, follow the setup steps above.

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
1. Edit `plugins/satva-brand-format/skills/satva-brand-format/SKILL.md`. Colours, fonts and the phone and email are in the `brand tokens` block of the helper module near the end of the file. If you changed the helper module, run `python tools/sync_module.py` so `satva_brand.py` matches.
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
