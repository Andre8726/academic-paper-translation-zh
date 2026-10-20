#!/usr/bin/env python3
"""Render and sanity-check a translated PDF.

Produces page PNGs, 2x2 contact sheets, and report.json with page count,
text presence, fonts, and suspicious glyphs. Visual inspection is still required.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def require_fitz():
    try:
        import fitz  # type: ignore
        return fitz
    except Exception as exc:
        sys.exit(f"PyMuPDF (fitz) is required: {exc}")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--expect-pages", type=int)
    args = ap.parse_args()

    fitz = require_fitz()
    try:
        from PIL import Image, ImageDraw
    except Exception as exc:
        sys.exit(f"Pillow is required for contact sheets: {exc}")

    pdf = Path(args.pdf).expanduser().resolve()
    out = Path(args.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(str(pdf))

    pages = []
    images = []
    page_paths = []
    for i, page in enumerate(doc):
        text = page.get_text("text")
        pix = page.get_pixmap(dpi=args.dpi, colorspace=fitz.csRGB)
        png = out / f"page-{i+1:02d}.png"
        pix.save(str(png))
        page_paths.append(png)
        pages.append({
            "page": i + 1,
            "width_pt": round(page.rect.width, 2),
            "height_pt": round(page.rect.height, 2),
            "chars": len(text),
            "images": len(page.get_images(full=True)),
            "replacement_chars": text.count("\ufffd"),
            "tofu_chars": text.count("\u25a1") + text.count("\u25a0"),
        })
        for img in page.get_images(full=True):
            images.append({"page": i + 1, "xref": img[0], "width_px": img[2], "height_px": img[3]})

    # Contact sheets: 2x2 pages per image.
    for start in range(0, len(page_paths), 4):
        group = page_paths[start:start + 4]
        ims = [Image.open(p) for p in group]
        w = max(im.width for im in ims)
        h = max(im.height for im in ims)
        scale = 0.8
        tw, th = int(w * scale), int(h * scale)
        sheet = Image.new("RGB", (tw * 2 + 30, th * 2 + 30), "white")
        dr = ImageDraw.Draw(sheet)
        for j, im in enumerate(ims):
            im = im.resize((int(im.width * scale), int(im.height * scale)))
            x = (j % 2) * tw + 10
            y = (j // 2) * th + 10
            sheet.paste(im, (x, y))
            dr.text((x + 8, y + 8), f"p{start + j + 1}", fill="red")
        sheet.save(out / f"contact-{(start // 4) + 1:02d}.png")

    # Fonts (unique base names).
    fonts = set()
    for page in doc:
        for f in page.get_fonts(full=True):
            fonts.add((f[3], f[4]))

    report = {
        "source": str(pdf),
        "page_count": doc.page_count,
        "expected_pages": args.expect_pages,
        "page_count_ok": args.expect_pages is None or doc.page_count == args.expect_pages,
        "total_chars": sum(p["chars"] for p in pages),
        "replacement_chars": sum(p["replacement_chars"] for p in pages),
        "tofu_chars": sum(p["tofu_chars"] for p in pages),
        "suspicious_pages": [
            p for p in pages
            if p["chars"] < 20 and p["images"] == 0
        ],
        "pages": pages,
        "images": images,
        "fonts": sorted([{"base": b, "name": n} for b, n in fonts], key=lambda x: x["name"]),
    }
    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({k: report[k] for k in (
        "page_count", "expected_pages", "page_count_ok", "total_chars",
        "replacement_chars", "tofu_chars", "suspicious_pages"
    )}, ensure_ascii=False, indent=2))
    print(f"wrote {out / 'report.json'} and contact sheets in {out}")


if __name__ == "__main__":
    main()
