# NotebookLM Math to Word

NotebookLM / Gemini / ChatGPT output exported through Google Docs keeps math as
literal LaTeX text (`$\pm 0.1\text{ mm}$`). This turns it into native, editable
Word equations. Three ways to use it:

| Way | Needs | Where |
|---|---|---|
| Web page: one HTML file, JS only, no backend | a browser | `site/index.html` → GitHub Pages |
| Python script | python3, lxml, pandoc | `skill/notebooklm-docx-math/scripts/docx_latex_to_word.py` |
| Claude Code skill | the above | `skill/notebooklm-docx-math/` |

## Script

```bash
pip install lxml && brew install pandoc    # or apt install pandoc
python3 skill/notebooklm-docx-math/scripts/docx_latex_to_word.py in.docx out.docx
```

## Skill

```bash
cp -r skill/notebooklm-docx-math ~/.claude/skills/
```

Then ask Claude: "make this NotebookLM docx readable".

## Web page

`site/index.html` is generated and self-contained (JSZip, temml and the
converter are inlined; no CDN, works offline). After editing `src/` or
`vendor/`, run `python3 build.py` and commit the result. The workflow in
`.github/workflows/pages.yml` publishes `site/` to GitHub Pages on push to `main`.

## How it works

Both engines find math spans per paragraph (across split Word runs), repair
Markdown damage typical of Google Docs exports (`\\text`, `\_`, eaten
subscripts), and splice `<m:oMath>` into `word/document.xml`.
Python uses pandoc for LaTeX→OMML; the page uses temml (LaTeX→MathML) plus a
small MathML→OMML mapper in `src/converter.js`.

Sample run (Google Doc about CFRP tooling): 220 equations, 0 left as text,
both engines.
