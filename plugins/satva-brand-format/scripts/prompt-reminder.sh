#!/usr/bin/env bash
# Runs on every prompt. If the request is about a document, deck, PDF, spreadsheet, page or email,
# remind Claude to load the Satva format. Prints nothing otherwise.
input="$(cat)"
if printf '%s' "$input" | grep -qiE 'document|\bdocs?\b|brd|frd|srs|\bsow\b|proposal|report|\bpdf|\bword\b|docx|\bdeck|slide|presentation|powerpoint|pptx|spreadsheet|excel|xlsx|\bsheet|tracker|letter|minutes|notes|one-pager|web ?page|artifact|\bhtml|e-?mail'; then
cat <<'JSON'
{"hookSpecificOutput":{"hookEventName":"UserPromptSubmit","additionalContext":"Satva house format: if this request creates or edits a document, report, PDF, deck, spreadsheet, Google Doc/Slides, web page or email draft, load the satva-brand-format skill (satva-brand-format:satva-brand-format) first and follow it."}}
JSON
fi
exit 0
