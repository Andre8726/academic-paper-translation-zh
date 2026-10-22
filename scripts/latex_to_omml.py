#!/usr/bin/env python3
"""Convert LaTeX equations to Office Math (OMML) XML files.

Input JSON:
  {
    "eq1": {"number": "1", "rows": ["\\frac{a}{b}=c"]},
    "eq2": {"number": "2", "rows": ["E=mc^2"]}
  }

Requires: latex2mathml, lxml, and Microsoft Word's mathml2omml.xsl
(default path on macOS). The output XML files are inserted by build_docx.py
through blocks of type "equation_omml".
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

DEFAULT_XSL = "/Applications/Microsoft Word.app/Contents/Resources/mathml2omml.xsl"


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--input", required=True, help="JSON map name -> {number, rows}")
    ap.add_argument("--out", required=True, help="output directory for OMML XML")
    ap.add_argument("--xsl", default=DEFAULT_XSL, help="Word mathml2omml.xsl path")
    args = ap.parse_args()

    try:
        from latex2mathml.converter import convert
        from lxml import etree
    except Exception as exc:
        sys.exit(
            "Missing dependency: %s\n"
            "Install with: python -m pip install --target \"$JOB/.mathdeps\" latex2mathml\n"
            "Then run with: PYTHONPATH=\"$JOB/.mathdeps\" python scripts/latex_to_omml.py ..." % exc
        )

    xsl_path = Path(args.xsl).expanduser()
    if not xsl_path.is_file():
        sys.exit(
            f"mathml2omml.xsl not found: {xsl_path}\n"
            "Pass --xsl /path/to/mathml2omml.xsl (Microsoft Word ships this file)."
        )
    xsl = etree.XSLT(etree.parse(str(xsl_path)))
    data = json.loads(Path(args.input).read_text(encoding="utf-8"))
    out = Path(args.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    mapping = {}

    for name, item in data.items():
        rows = item.get("rows") or ([item["tex"]] if item.get("tex") else [])
        if not rows:
            raise SystemExit(f"{name}: provide rows[] or tex")
        files = []
        for i, tex in enumerate(rows):
            tex = tex.strip()
            if tex.startswith("$") and tex.endswith("$"):
                tex = tex[1:-1]
            mathml = convert(tex)
            dom = etree.fromstring(mathml.encode("utf-8"))
            omml = str(xsl(dom))
            fn = f"{name}_{i}.xml"
            (out / fn).write_text(omml, encoding="utf-8")
            files.append(str(out / fn))
        mapping[name] = {"number": str(item.get("number", "")), "files": files}
        print(f"{name}: {len(files)} row(s)")

    (out / "omml_map.json").write_text(
        json.dumps(mapping, ensure_ascii=False, indent=1), encoding="utf-8"
    )
    print(f"wrote {out / 'omml_map.json'} ({len(mapping)} equations)")


if __name__ == "__main__":
    main()
