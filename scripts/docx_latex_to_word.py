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


def normalize_tex(tex: str) -> str:
    """^\\circ is the degree sign; a bare \\circ renders as the small-circle operator."""
    return re.sub(r"\^\s*\{\s*\\circ\s*\}|\^\s*\\circ(?![A-Za-z])", "\u00b0", tex)


def split_csv(line):
    """Split one CSV-ish line; commas inside "quotes" or $math$ do not split. None if unbalanced."""
    cells, cur, q, math, started, i = [], "", False, False, False, 0
    while i < len(line):
        ch = line[i]
        if q:
            if ch == '"':
                if line[i + 1:i + 2] == '"':
                    cur += '"'
                    i += 1
                else:
                    q = False
            else:
                cur += ch
        elif ch == '"' and not started:
            q = started = True
        elif ch == "$":
            math = not math
            cur += ch
            started = True
        elif ch == "," and not math:
            cells.append(cur.strip())
            cur, started = "", False
        else:
            cur += ch
            started = started or bool(ch.strip())
        i += 1
    if q or math:
        return None
    cells.append(cur.strip())
    return cells


def _csv_candidate(p):
    ch = [c for c in p if c.tag != f"{{{W}}}pPr"]
    if not ch or not all(is_text_run(c) for c in ch):
        return None
    ppr = p.find("w:pPr", NS)
    if ppr is not None:
        if ppr.find("w:numPr", NS) is not None:
            return None
        st = ppr.find("w:pStyle", NS)
        if st is not None and re.match(r"(Heading|Title|ListParagraph)", st.get(f"{{{W}}}val", ""), re.I):
            return None
    cells = split_csv("".join("".join(t.text or "" for t in r.findall("w:t", NS)) for r in ch))
    return (p, cells, ch[0]) if cells and len(cells) >= 2 else None


def _w(tag, **attrs):
    e = etree.Element(f"{{{W}}}{tag}")
    for k, v in attrs.items():
        e.set(f"{{{W}}}{k}", v)
    return e


def _build_table(rows):
    tbl, pr = _w("tbl"), _w("tblPr")
    pr.append(_w("tblW", w="5000", type="pct"))
    b = _w("tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b.append(_w(side, val="single", sz="4", space="0", color="999999"))
    pr.append(b)
    cm = _w("tblCellMar")
    for side in ("left", "right"):
        cm.append(_w(side, w="80", type="dxa"))
    pr.append(cm)
    tbl.append(pr)
    grid = _w("tblGrid")
    for _ in rows[0][1]:
        grid.append(_w("gridCol"))
    tbl.append(grid)
    for ri, (_p, cells, src) in enumerate(rows):
        tr = _w("tr")
        if ri == 0:
            trp = _w("trPr")
            trp.append(_w("tblHeader"))
            tr.append(trp)
        for text in cells:
            tc, p, r = _w("tc"), _w("p"), text_run(src, text)
            if ri == 0:
                rpr = r.find("w:rPr", NS)
                if rpr is None:
                    rpr = _w("rPr")
                    r.insert(0, rpr)
                rpr.append(_w("b"))
            p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)
    return tbl


def convert_tables(root):
    """Runs of consecutive CSV-looking paragraphs become real Word tables. Returns count."""
    count = 0
    parents = {p.getparent() for p in root.iter(f"{{{W}}}p")}
    prose = lambda g: any(len(c) > 80 for c in g[1]) or bool(re.search(r"[:.]$", g[1][-1]))
    for parent in parents:
        group = []

        def flush(nxt=None):
            nonlocal group, count
            while group and prose(group[0]):
                group.pop(0)
            n = len(group[0][1]) if group else 0
            if len(group) >= 2 and (n >= 3 or (len(group) >= 3 and all(len(c) < 60 for g in group for c in g[1]))):
                parent.insert(parent.index(group[0][0]), _build_table(group))
                for g in group:
                    parent.remove(g[0])
                count += 1
            group = [nxt] if nxt else []

        for ch in list(parent):
            if ch.tag == f"{{{W}}}p":
                c = _csv_candidate(ch)
                if c and (not group or len(group[0][1]) == len(c[1])):
                    group.append(c)
                else:
                    flush(c)
            else:
                flush()
        flush()
    return count


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


def convert(src, dst, fix=True, tables=True):
    if not shutil.which("pandoc"):
        sys.exit("pandoc not found. Install: brew install pandoc | apt install pandoc")
    with zipfile.ZipFile(src) as zin:
        parts = {n: zin.read(n) for n in zin.namelist()}
        infos = {i.filename: i for i in zin.infolist()}
    trees, all_jobs, exprs, ntables = {}, [], [], 0
    for name in parts:
        if PART_RE.match(name):
            root = etree.fromstring(parts[name])
            ntab = convert_tables(root) if tables else 0
            ntables += ntab
            jobs = collect_matches(root, fix)
            if ntab and not jobs:
                trees[name] = (root, [])
            if jobs:
                trees[name] = (root, jobs)
                for _p, _g, _t, ms in jobs:
                    for m in ms:
                        tex = next(v for k, v in m.groupdict().items() if v is not None)
                        exprs.append(normalize_tex(fix_markdown_damage(tex) if fix else tex))
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
    return done, skipped, ntables


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--no-tables", action="store_true", help="keep comma-separated text as text")
    ap.add_argument("--no-fix-markdown", action="store_true", help="do not repair \\\\text / \\_ damage")
    a = ap.parse_args()
    src = Path(a.input)
    dst = Path(a.output) if a.output else src.with_name(src.stem + "_math.docx")
    done, skipped, nt = convert(src, dst, fix=not a.no_fix_markdown, tables=not a.no_tables)
    print(f"{dst}: {done} equations converted, {skipped} left as text, {nt} tables built")


if __name__ == "__main__":
    main()
