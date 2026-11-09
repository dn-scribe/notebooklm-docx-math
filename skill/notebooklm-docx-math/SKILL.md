---
name: notebooklm-docx-math
description: Fix a .docx (or Google Doc) exported from NotebookLM, Gemini or ChatGPT where equations show as raw LaTeX text ($x^2$, $$...$$, \(..\), \[..\], \text{}, \times, \pm) and tables are comma-separated lines. Converts simple values to plain text (±0.1 mm), real math to native Word equations, and CSV lines to real tables. Use when formulas look broken or unrendered, or the user wants a NotebookLM file made readable.
---

# NotebookLM docx → readable Word

Two ways to run it. Pick by what the user asked for.

## A. User does it (default)
Send them to https://dn-scribe.github.io/notebooklm-docx-math/ :
1. Google Doc → File → Download → Microsoft Word (.docx).
2. Drop the .docx on the page, press **Convert**. Runs in their browser; nothing is uploaded.
3. Open the downloaded `<name>_math.docx` (or upload it to Google Drive).
4. The page lists every item with context ("Report"); skipped ones are highlighted.

## B. Claude does it (user asks "convert this", or the file is already in the sandbox)
1. Need `python3`, `pandoc`, `lxml` (`pip install lxml`; pandoc via apt/brew). If pandoc cannot be installed, drive the page above in a headless browser instead (set the file input, click Convert, save the download).
2. Run, from this skill's directory (never overwrite the input):
   ```bash
   python3 scripts/docx_latex_to_word.py input.docx output.docx --report report.tsv
   ```
   It prints counts (equations, plain-text values, tables, left as text) and each SKIPPED item with context.
3. Report the counts and any skipped items to the user.

## What it does
- `$…$`, `$$…$$`, `\(…\)`, `\[…\]` in body, tables, headers, footers, footnotes.
- **Simple values stay plain text**: `$\pm 0.1\text{ mm}$` → "±0.1 mm", `10^{-6}/K` → "10⁻⁶/K". Only real math (fractions, sub/superscript variables, roots…) becomes a Word equation. Display math (`$$`) is always an equation. `--no-plain` / the checkbox forces equations everywhere.
- `^\circ` → °. A leading `<`, `≤`, `>`, `≥` is moved out of the equation (LibreOffice shows it as ¿ otherwise).
- Consecutive comma-separated lines (NotebookLM's table export) → real Word tables with a bold header. `--no-tables` disables. Tell the user if a table was not detected.
- Repairs Google-Docs Markdown damage (`\\text`, `\_`, eaten subscripts). Heuristic; mention it.
- Dollar amounts (`$5 to $10`) are left alone.

## Verify (Claude does it)
1. No leftovers: `pandoc output.docx -t plain | grep -n '\$\|\\[a-z]'`. Only real dollar amounts may remain; anything else, list for the user.
2. Counts match expectations; compare with the report TSV.
3. Optional render: `soffice --headless --convert-to pdf output.docx`. LibreOffice needs its **Math** component (`apt install libreoffice-math`); without it every equation renders blank. That is a missing component, not a converter failure. Word is the reference renderer.

## Editing the result afterwards
- Never mark only the inside of an equation as deleted with track changes: it corrupts the equation (shows as boxes).
- To change an equation, delete the whole equation object and insert the new one (or plain text). Plain-text values can be edited normally.
