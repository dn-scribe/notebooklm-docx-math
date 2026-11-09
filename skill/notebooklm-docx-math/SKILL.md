---
name: notebooklm-docx-math
description: Convert a .docx whose equations are raw LaTeX text ($x^2$, $$...$$, \(..\), \[..\]) — typical of NotebookLM, Gemini or ChatGPT output pasted or exported via Google Docs — into a .docx with native, editable Word equations. Use when the user has a Word/Google-Doc export full of dollar-sign math, backslash commands like \text{} \times \pm, or asks to make NotebookLM files readable.
---

# NotebookLM docx math → Word equations

## When
User hands over a `.docx` (or a Google Doc link) where math shows as literal
`$\pm 0.1\text{ mm}$`-style text and wants it readable.

## Steps
1. Get a local `.docx`. For a Google Doc, export as
   `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
   (Drive `download_file_content`; decode the base64 `content` field to a file).
2. Check deps: `pandoc -v` and `python3 -c "import lxml"`.
   Install if missing: `brew install pandoc` / `apt install pandoc`, `pip install lxml`.
3. Run (resolve the script relative to this skill's directory):
   ```bash
   python3 scripts/docx_latex_to_word.py input.docx [output.docx]
   ```
   Default output is `<input>_math.docx` next to the input. Never overwrite the input.
4. Read the printed counts. `N left as text` means pandoc could not parse that
   expression; list them (grep the output text for `$`) and tell the user.
5. Verify: `pandoc output.docx -t plain | grep -c '\$'` should be near zero
   (stray `$` in currency like `$0.35` is expected and intentionally untouched).
6. Report: path of output, converted/skipped counts, any repairs applied.

## What the script does
- Finds `$…$`, `$$…$$`, `\(…\)`, `\[…\]` per paragraph, even when Word split the
  text over several runs; keeps surrounding run formatting.
- Display math that is the whole paragraph becomes a centered `m:oMathPara`;
  everything else is inline `m:oMath`.
- Skips currency (`$5 to $10`): a closing `$` followed by a digit is not math.
- Repairs Markdown damage from Google Docs exports (disable with `--no-fix-markdown`):
  `\\text`→`\text`, `\\%`→`\%`, `\_`→`_`, `\*`→`*`, and `\text{A}{\text{B}}`→`\text{A}_{\text{B}}`
  (subscript underscore eaten as italics). Mention this to the user: it is a heuristic.
- Processes body, tables, headers, footers, footnotes, endnotes.

## Limits
- Needs `pandoc`. No pandoc → use the browser version (`site/index.html`, no install).
- Only equations delimited as above. Bare LaTeX with no delimiters is untouched.
- Does not render to check visually; open in Word and spot-check one or two
  display equations if the user cares.
