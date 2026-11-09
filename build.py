#!/usr/bin/env python3
"""Build site/index.html: one self-contained file (JSZip + temml + converter inlined).
Run after editing src/ or vendor/:  python3 build.py"""
from pathlib import Path

root = Path(__file__).parent
page = (root / "src/index.template.html").read_text(encoding="utf-8")
for key, path in [("jszip", "vendor/jszip.min.js"), ("temml", "vendor/temml.min.js"), ("converter", "src/converter.js")]:
    code = (root / path).read_text(encoding="utf-8").replace("</script", "<\\/script")
    page = page.replace(f"/*@@{key}@@*/", code)
out = root / "site/index.html"
out.write_text(page, encoding="utf-8")
print(f"{out} ({out.stat().st_size // 1024} KB)")
