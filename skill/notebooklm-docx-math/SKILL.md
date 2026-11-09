---
name: notebooklm-docx-math
description: Fix a .docx (or Google Doc) where equations show as raw LaTeX text — $x^2$, $$...$$, \(..\), \[..\], \text{}, \times, \pm — typical of NotebookLM, Gemini or ChatGPT output exported to Word or Google Docs. Sends the user to a free in-browser converter that turns them into native, editable Word equations. Use when the user says formulas/equations look broken, unrendered, or full of dollar signs and backslashes, or wants a NotebookLM file made readable.
---

# NotebookLM docx math → Word equations

The conversion runs in a web page, not here. Do not try to convert the file
yourself or install anything.

Converter: https://dn-scribe.github.io/notebooklm-docx-math/

## What to tell the user
1. If the file is a Google Doc: File → Download → Microsoft Word (.docx).
2. Open the converter link above.
3. Drop the .docx on the page (or click to choose it) and press **Convert**.
   Everything runs in their browser; nothing is uploaded.
4. A new file `<name>_math.docx` downloads. Open it in Word (or upload it back
   to Google Drive / Google Docs, which keeps the equations).
5. The page reports "N equations converted" and, if any, "M left as text".
   Left-as-text items are LaTeX the converter could not parse; offer to help
   fix those by hand.

## Notes
- Handles `$…$`, `$$…$$`, `\(…\)`, `\[…\]` in body, tables, headers, footers, footnotes.
- Dollar amounts like `$5 to $10` are left alone.
- "Repair Markdown damage" (on by default) fixes `\\text`, `\_` and eaten
  subscripts from Google Docs exports. If equations look wrong, they can untick it and retry.
- Source and an offline Python version: https://github.com/dn-scribe/notebooklm-docx-math
