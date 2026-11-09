/* LaTeX-in-docx -> native Word equations, fully client side.
 * LaTeX --(temml)--> MathML --(this file)--> OMML, spliced into word/document.xml.
 * Deps are injected so the same code runs in the browser and under node tests. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.DocxMath = factory();
})(typeof self !== "undefined" ? self : this, function () {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const MML = "http://www.w3.org/1998/Math/MathML";
  const PART_RE = /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/;
  const MATH_RE =
    /\$\$(?<d1>[\s\S]+?)\$\$|\\\[(?<d2>[\s\S]+?)\\\]|\\\((?<i2>[\s\S]+?)\\\)|(?<![\\$\w])\$(?![\s$])(?<i1>[^$\n]+?)(?<![\s\\])\$(?![\d$])/g;

  const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  /* ^\circ is the degree sign; \circ alone renders as the small-circle operator. */
  function normalizeTex(tex) {
    return tex.replace(/\^\s*\{\s*\\circ\s*\}|\^\s*\\circ(?![A-Za-z])/g, "\u00B0");
  }

  /* Undo Markdown mangling typical of Google-Docs exports of LLM output. */
  function fixMarkdownDamage(tex) {
    tex = tex.replace(/\\\\(?=[A-Za-z%&#])/g, "\\");
    tex = tex.replace(/\\_/g, "_").replace(/\\\*/g, "*");
    return tex.replace(/(\\text\{[^{}]*\})(\{\\text\{)/g, "$1_$2");
  }

  /* ---------- MathML -> OMML string ---------- */
  const ACCENTS = { "^": "̂", "ˆ": "̂", "¯": "̄", "‾": "̄", "~": "̃", "˜": "̃", "→": "⃗", "˙": "̇", "¨": "̈", "ˇ": "̌" };

  function kids(el) { return Array.from(el.childNodes).filter((n) => n.nodeType === 1); }
  function seq(el) { return kids(el).map(conv).join(""); }
  function wrap(tag, el) { return `<m:${tag}>${el}</m:${tag}>`; }
  function run(text, props) {
    return `<m:r>${props ? `<m:rPr>${props}</m:rPr>` : ""}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
  }

  function conv(el) {
    const k = kids(el), n = el.localName;
    const c = (i) => (k[i] ? conv(k[i]) : "");
    switch (n) {
      case "math": case "mrow": case "mstyle": case "mpadded": case "semantics": case "menclose": case "merror":
        return k.filter((x) => x.localName !== "annotation").map(conv).join("");
      case "mi": {
        const t = el.textContent;
        return run(t, t.length > 1 || el.getAttribute("mathvariant") === "normal" ? '<m:sty m:val="p"/>' : "");
      }
      case "mn": return run(el.textContent, "");
      case "mo": return run(el.textContent, "");
      case "mtext": return run(el.textContent, "<m:nor/>");
      case "mspace": {
        const w = parseFloat(el.getAttribute("width") || "0");
        return w >= 0.9 ? run("\u00A0", "") : "";
      }
      case "mfrac": return wrap("f", wrap("num", c(0)) + wrap("den", c(1)));
      case "msqrt": return wrap("rad", '<m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>' + wrap("e", seq(el)));
      case "mroot": return wrap("rad", wrap("deg", c(1)) + wrap("e", c(0)));
      case "msup": return wrap("sSup", wrap("e", c(0)) + wrap("sup", c(1)));
      case "msub": return wrap("sSub", wrap("e", c(0)) + wrap("sub", c(1)));
      case "msubsup": return wrap("sSubSup", wrap("e", c(0)) + wrap("sub", c(1)) + wrap("sup", c(2)));
      case "mover": {
        const a = k[1] && ACCENTS[k[1].textContent.trim()];
        if (a) return wrap("acc", `<m:accPr><m:chr m:val="${a}"/></m:accPr>` + wrap("e", c(0)));
        return wrap("limUpp", wrap("e", c(0)) + wrap("lim", c(1)));
      }
      case "munder": return wrap("limLow", wrap("e", c(0)) + wrap("lim", c(1)));
      case "munderover":
        return wrap("limUpp", wrap("e", wrap("limLow", wrap("e", c(0)) + wrap("lim", c(1)))) + wrap("lim", c(2)));
      case "mphantom": return wrap("phant", wrap("e", seq(el)));
      case "mtable":
        return wrap("m", k.map((tr) => wrap("mr", kids(tr).map((td) => wrap("e", seq(td))).join(""))).join(""));
      case "mtr": case "mtd": return seq(el);
      default: return seq(el) || (k.length ? "" : run(el.textContent, ""));
    }
  }

  function latexToOmml(tex, display, deps) {
    let mml;
    try {
      mml = deps.temml.renderToString(tex, { displayMode: display, throwOnError: true, xml: true });
    } catch (e) { return null; }
    const doc = new deps.DOMParser().parseFromString(mml, "application/xml");
    const root = doc.documentElement;
    if (!root || root.localName === "parsererror" || doc.getElementsByTagName("parsererror").length) return null;
    return conv(root);
  }

  /* ---------- docx walking ---------- */
  const isTextRun = (el, ns) =>
    el.namespaceURI === W && el.localName === "r" &&
    kids(el).some((x) => x.localName !== "rPr") &&
    kids(el).every((x) => x.localName === "rPr" || x.localName === "t");
  const runText = (r) => kids(r).filter((x) => x.localName === "t").map((t) => t.textContent).join("");

  function makeRun(doc, srcRun, text) {
    const r = doc.createElementNS(W, "w:r");
    const rpr = kids(srcRun).find((x) => x.localName === "rPr");
    if (rpr) r.appendChild(rpr.cloneNode(true));
    const t = doc.createElementNS(W, "w:t");
    t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
    t.textContent = text;
    r.appendChild(t);
    return r;
  }

  /* ---------- CSV-as-text paragraphs -> real Word tables ---------- */
  function splitCsv(line) {
    const cells = []; let cur = "", q = false, math = false, started = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) {
        if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else q = false; }
        else cur += ch;
      } else if (ch === '"' && !started) { q = true; started = true; }
      else if (ch === "$") { math = !math; cur += ch; started = true; }
      else if (ch === "," && !math) { cells.push(cur.trim()); cur = ""; started = false; }
      else { cur += ch; if (ch.trim()) started = true; }
    }
    if (q || math) return null;
    cells.push(cur.trim());
    return cells;
  }

  function csvCandidate(p) {
    const ch = kids(p).filter((x) => x.localName !== "pPr");
    if (!ch.length || !ch.every((x) => isTextRun(x))) return null;
    const ppr = kids(p).find((x) => x.localName === "pPr");
    if (ppr) {
      const st = kids(ppr).find((x) => x.localName === "pStyle");
      if (kids(ppr).some((x) => x.localName === "numPr")) return null;
      if (st && /^(Heading|Title|ListParagraph)/i.test(st.getAttributeNS(W, "val") || "")) return null;
    }
    const cells = splitCsv(ch.map(runText).join(""));
    return cells && cells.length >= 2 ? { p, cells, run: ch[0] } : null;
  }

  function buildTable(doc, rows) {
    const el = (n, attrs = {}) => {
      const e = doc.createElementNS(W, "w:" + n);
      for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, "w:" + k, v);
      return e;
    };
    const tbl = el("tbl"), pr = el("tblPr");
    pr.appendChild(el("tblW", { w: "5000", type: "pct" }));
    const b = el("tblBorders");
    for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"])
      b.appendChild(el(side, { val: "single", sz: "4", space: "0", color: "999999" }));
    pr.appendChild(b);
    const cm = el("tblCellMar");
    for (const side of ["left", "right"]) cm.appendChild(el(side, { w: "80", type: "dxa" }));
    pr.appendChild(cm);
    tbl.appendChild(pr);
    const grid = el("tblGrid");
    for (let i = 0; i < rows[0].cells.length; i++) grid.appendChild(el("gridCol"));
    tbl.appendChild(grid);
    rows.forEach((row, ri) => {
      const tr = el("tr");
      if (ri === 0) { const trp = el("trPr"); trp.appendChild(el("tblHeader")); tr.appendChild(trp); }
      for (const text of row.cells) {
        const tc = el("tc"), p = el("p"), r = makeRun(doc, row.run, text);
        if (ri === 0) {
          let rpr = kids(r).find((x) => x.localName === "rPr");
          if (!rpr) { rpr = el("rPr"); r.insertBefore(rpr, r.firstChild); }
          rpr.appendChild(el("b"));
        }
        p.appendChild(r); tc.appendChild(p); tr.appendChild(tc);
      }
      tbl.appendChild(tr);
    });
    return tbl;
  }

  function convertTables(doc, stats) {
    const parents = new Set();
    for (const p of Array.from(doc.getElementsByTagNameNS(W, "p"))) parents.add(p.parentNode);
    for (const parent of parents) {
      let group = [];
      // intro sentences with the same comma count as the table are not header rows
      const prose = (g) => g.cells.some((c) => c.length > 80) || /[:.]$/.test(g.cells[g.cells.length - 1]);
      const flush = (next) => {
        while (group.length && prose(group[0])) group.shift();
        const n = group.length ? group[0].cells.length : 0;
        const ok = group.length >= 2 &&
          (n >= 3 || (group.length >= 3 && group.every((g) => g.cells.every((c) => c.length < 60))));
        if (ok) {
          const t = buildTable(doc, group);
          parent.insertBefore(t, group[0].p);
          // Word needs a paragraph after a table that ends a cell/body
          for (const g of group) parent.removeChild(g.p);
          stats.tables++;
        }
        group = [];
        if (next) group.push(next);
      };
      for (const ch of kids(parent)) {
        if (ch.namespaceURI === W && ch.localName === "p") {
          const c = csvCandidate(ch);
          if (c && (!group.length || group[0].cells.length === c.cells.length)) group.push(c);
          else flush(c);
        } else flush(null);
      }
      flush(null);
    }
  }

  function processPart(xml, deps, opts, stats) {
    const doc = new deps.DOMParser().parseFromString(xml, "application/xml");
    const tablesBefore = stats.tables;
    if (opts.tables) convertTables(doc, stats);
    const paras = Array.from(doc.getElementsByTagNameNS(W, "p"));
    const jobs = [];
    for (const p of paras) {
      const groups = []; let cur = [];
      for (const ch of kids(p)) {
        if (isTextRun(ch)) cur.push(ch); else { if (cur.length) groups.push(cur); cur = []; }
      }
      if (cur.length) groups.push(cur);
      for (const g of groups) {
        const text = g.map(runText).join("");
        if (!/\$|\\\(|\\\[/.test(text)) continue;
        const ms = Array.from(text.matchAll(MATH_RE));
        if (ms.length) jobs.push({ p, g, text, ms });
      }
    }
    for (const { p, g, text, ms } of jobs) {
      const owner = [];
      for (const r of g) for (let i = runText(r).length; i--;) owner.push(r);
      const anchor = g[0], out = [];
      let last = 0;
      for (const m of ms) {
        const gr = m.groups;
        const tex0 = gr.d1 ?? gr.d2 ?? gr.i2 ?? gr.i1;
        const display = gr.d1 !== undefined || gr.d2 !== undefined;
        const omml = latexToOmml(normalizeTex(opts.fix ? fixMarkdownDamage(tex0) : tex0), display, deps);
        if (m.index > last) out.push(makeRun(doc, owner[last], text.slice(last, m.index)));
        if (omml === null) {
          out.push(makeRun(doc, owner[m.index], m[0])); stats.skipped++;
        } else {
          const whole = display && m.index === 0 && m.index + m[0].length === text.length;
          const inner = `<m:oMath>${omml}</m:oMath>`;
          const frag = new deps.DOMParser().parseFromString(
            `<x xmlns:m="${M}" xmlns:w="${W}">${whole ? `<m:oMathPara>${inner}</m:oMathPara>` : inner}</x>`,
            "application/xml");
          out.push(doc.importNode(frag.documentElement.firstChild, true)); stats.done++;
        }
        last = m.index + m[0].length;
      }
      if (last < text.length) out.push(makeRun(doc, owner[last], text.slice(last)));
      for (const el of out) p.insertBefore(el, anchor);
      for (const r of g) p.removeChild(r);
    }
    return jobs.length || stats.tables > tablesBefore ? new deps.XMLSerializer().serializeToString(doc) : null;
  }

  /* deps: { JSZip, temml, DOMParser, XMLSerializer }; input: ArrayBuffer/Uint8Array */
  async function convert(input, deps, opts = {}) {
    opts = { fix: true, tables: true, ...opts };
    const zip = await deps.JSZip.loadAsync(input);
    const stats = { done: 0, skipped: 0, tables: 0 };
    for (const name of Object.keys(zip.files)) {
      if (!PART_RE.test(name)) continue;
      const out = processPart(await zip.file(name).async("string"), deps, opts, stats);
      if (out !== null) zip.file(name, out);
    }
    const blob = await zip.generateAsync({
      type: deps.outType || "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
    });
    return { blob, ...stats };
  }

  return { convert, fixMarkdownDamage, normalizeTex, splitCsv, latexToOmml };
});
