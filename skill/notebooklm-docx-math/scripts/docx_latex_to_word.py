#!/usr/bin/env python3
"""Convert raw LaTeX in a .docx ($x^2$, $$...$$, \\(..\\), \\[..\\]) into native Word equations.

NotebookLM / Gemini / ChatGPT exports to Google Docs leave math as literal
text such as ``$\\pm 0.1\\text{ mm}$``. This rewrites those spans as OMML
(``<m:oMath>``) so Word renders them as real, editable equations.

Needs: python3, lxml, pandoc (LaTeX -> OMML engine).
Usage: docx_latex_to_word.py in.docx [out.docx] [--no-fix-markdown]
"""
import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
NS = {"w": W, "m": M}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Parts that hold body text. Order irrelevant.
PART_RE = re.compile(r"^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")

# $$..$$ | \[..\] | \(..\) | $..$   (display flag decided by group)
MATH_RE = re.compile(
    r"\$\$(?P<d1>.+?)\$\$"
    r"|\\\[(?P<d2>.+?)\\\]"
    r"|\\\((?P<i2>.+?)\\\)"
    r"|(?<![\\$\w])\$(?![\s$])(?P<i1>[^$\n]+?)(?<![\s\\])\$(?![\d$])",
    re.S,
)


def fix_markdown_damage(tex: str) -> str:
    """Undo Markdown mangling common in Google-Docs exports of LLM output."""
    tex = re.sub(r"\\\\(?=[A-Za-z%&#])", r"\\", tex)  # \\text -> \text
    tex = tex.replace("\\_", "_").replace("\\*", "*")
    # subscripts eaten as italics: \text{CTE}{\text{Part}} -> \text{CTE}_{\text{Part}}
    tex = re.sub(r"(\\text\{[^{}]*\})(\{\\text\{)", r"\1_\2", tex)
    return tex


def latex_to_omml(exprs):
    """Batch-convert LaTeX strings with one pandoc call. Returns list of oMath
    elements (None where pandoc could not parse the expression)."""
    if not exprs:
        return []
    md = "\n\n".join("$" + " ".join(e.split()) + "$" for e in exprs) + "\n"
    with tempfile.TemporaryDirectory() as td:
        src, out = Path(td, "m.md"), Path(td, "m.docx")
        src.write_text(md, encoding="utf-8")
        subprocess.run(
            ["pandoc", "-f", "markdown+tex_math_dollars", "-t", "docx", "-o", str(out), str(src)],
            check=True, capture_output=True,
        )
        with zipfile.ZipFile(out) as z:
            root = etree.fromstring(z.read("word/document.xml"))
    paras = root.findall(".//w:body/w:p", NS)
    if len(paras) != len(exprs):
        raise RuntimeError(f"pandoc returned {len(paras)} paragraphs for {len(exprs)} equations")
    res = []
    for p in paras:
        om = p.findall(".//m:oMath", NS)
        # exactly one oMath and no stray literal text => parsed OK
        res.append(om[0] if len(om) == 1 and not p.findall("./w:r/w:t", NS) else None)
    return res


def text_run(src_run, text):
    """Clone a run's properties, give it new text."""
    r = etree.Element(f"{{{W}}}r")
    rpr = src_run.find("w:rPr", NS)
    if rpr is not None:
        r.append(etree.fromstring(etree.tostring(rpr)))
    t = etree.SubElement(r, f"{{{W}}}t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def is_text_run(el):
    if el.tag != f"{{{W}}}r":
        return False
    kids = [k for k in el if k.tag != f"{{{W}}}rPr"]
    return bool(kids) and all(k.tag == f"{{{W}}}t" for k in kids)


def collect_matches(root, fix):
    """Pass 1: find math spans per paragraph. Returns list of (para, groups, matches)."""
    jobs = []
    for p in root.iter(f"{{{W}}}p"):
        groups, cur = [], []
        for ch in p:
            if is_text_run(ch):
                cur.append(ch)
            else:
                if cur:
                    groups.append(cur)
                cur = []
        if cur:
            groups.append(cur)
        for g in groups:
            text = "".join("".join(t.text or "" for t in r.findall("w:t", NS)) for r in g)
            if "$" not in text and "\\(" not in text and "\\[" not in text:
                continue
            ms = list(MATH_RE.finditer(text))
            if ms:
                jobs.append((p, g, text, ms))
    return jobs


def rewrite(jobs, omml_iter, fix):
    done = skipped = 0
    for p, g, text, ms in jobs:
        # char index -> source run
        owner = []
        for r in g:
            owner += [r] * len("".join(t.text or "" for t in r.findall("w:t", NS)))
        anchor = g[0]
        pos = p.index(anchor)
        for r in g:
            p.remove(r)
        out, last = [], 0
        for m in ms:
            display = bool(m.group("d1") or m.group("d2"))
            om = next(omml_iter)
            if m.start() > last:
                out.append(text_run(owner[last], text[last:m.start()]))
            if om is None:
                out.append(text_run(owner[m.start()], m.group(0)))
                skipped += 1
            else:
                om = etree.fromstring(etree.tostring(om))
                if display and m.start() == 0 and m.end() == len(text):
                    wrap = etree.Element(f"{{{M}}}oMathPara", nsmap={"m": M})
                    wrap.append(om)
                    out.append(wrap)
                else:
                    out.append(om)
                done += 1
            last = m.end()
        if last < len(text):
            out.append(text_run(owner[last], text[last:]))
        for i, el in enumerate(out):
            p.insert(pos + i, el)
    return done, skipped


def convert(src, dst, fix=True):
    if not shutil.which("pandoc"):
        sys.exit("pandoc not found. Install: brew install pandoc | apt install pandoc")
    with zipfile.ZipFile(src) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
        infos = {i.filename: i for i in zin.infolist()}
    trees, all_jobs, exprs = {}, [], []
    for name in parts:
        if PART_RE.match(name):
            root = etree.fromstring(parts[name])
            jobs = collect_matches(root, fix)
            if jobs:
                trees[name] = (root, jobs)
                for _p, _g, _t, ms in jobs:
                    for m in ms:
                        tex = next(v for k, v in m.groupdict().items() if v is not None)
                        exprs.append(fix_markdown_damage(tex) if fix else tex)
    omml = iter(latex_to_omml(exprs))
    done = skipped = 0
    for name, (root, jobs) in trees.items():
        d, s = rewrite(jobs, omml, fix)
        done += d
        skipped += s
        parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(infos[name], data)
    return done, skipped


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--no-fix-markdown", action="store_true", help="do not repair \\\\text / \\_ damage")
    a = ap.parse_args()
    src = Path(a.input)
    dst = Path(a.output) if a.output else src.with_name(src.stem + "_math.docx")
    done, skipped = convert(src, dst, fix=not a.no_fix_markdown)
    print(f"{dst}: {done} equations converted, {skipped} left as text")


if __name__ == "__main__":
    main()
