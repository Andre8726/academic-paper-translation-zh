#!/usr/bin/env python3
"""Inventory and extract text, blocks, embedded images, and page previews from a PDF.

Requires PyMuPDF (import name: fitz). Run with a Python that can import it.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def require_fitz():
    try:
        import fitz  # type: ignore
        return fitz
    except Exception as exc:  # pragma: no cover - environment dependent
        sys.exit(
            "PyMuPDF (fitz) is not importable: %s\n"
            "macOS workaround:\n"
            "  /usr/bin/python3 -m pip install --target \"$JOB/.pymupdf\" pymupdf==1.24.14\n"
            "  PYTHONPATH=\"$JOB/.pymupdf\" /usr/bin/python3 scripts/extract_pdf.py ...\n"
            "or run scripts/bootstrap_pymupdf.sh" % exc
        )


def parse_pages(expr: str, page_count: int) -> list[int]:
    if not expr or expr.strip().lower() == "all":
        return list(range(1, page_count + 1))
    pages: set[int] = set()
    for part in expr.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start, end = int(a), int(b)
            if start > end:
                start, end = end, start
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    invalid = [p for p in pages if p < 1 or p > page_count]
    if invalid:
        raise SystemExit(f"Page(s) out of range 1..{page_count}: {invalid}")
    return sorted(pages)


def page_dict(page, fitz) -> dict:
    blocks = []
    for b in page.get_text("blocks", sort=True):
        x0, y0, x1, y1, text, block_no, block_type = b[:7]
        blocks.append(
            {
                "bbox": [round(x0, 2), round(y0, 2), round(x1, 2), round(y1, 2)],
                "text": text.strip(),
                "block_no": block_no,
                "block_type": block_type,
            }
        )
    return {
        "number": page.number + 1,
        "width_pt": round(page.rect.width, 2),
        "height_pt": round(page.rect.height, 2),
        "rotation": page.rotation,
        "text": page.get_text("text"),
        "blocks": blocks,
        "images": [],
    }


def inventory(args) -> None:
    fitz = require_fitz()
    pdf = Path(args.pdf).expanduser().resolve()
    if not pdf.is_file():
        raise SystemExit(f"PDF not found: {pdf}")
    out = Path(args.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf))
    if doc.needs_pass:
        raise SystemExit("PDF is password-protected; decrypt it first.")
    pages = []
    for page in doc:
        images = []
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                rects = [list(map(lambda v: round(v, 2), r)) for r in page.get_image_rects(xref)]
            except Exception:
                rects = []
            images.append(
                {
                    "xref": xref,
                    "width_px": img[2],
                    "height_px": img[3],
                    "bpc": img[4],
                    "colorspace": img[5],
                    "rects": rects,
                }
            )
        pages.append(
            {
                "number": page.number + 1,
                "width_pt": round(page.rect.width, 2),
                "height_pt": round(page.rect.height, 2),
                "chars": len(page.get_text("text")),
                "images": images,
            }
        )
    total = sum(p["chars"] for p in pages)
    report = {
        "source": str(pdf),
        "page_count": doc.page_count,
        "total_chars": total,
        "likely_scanned": total < max(200, doc.page_count * 120),
        "pages": pages,
    }
    (out / "inventory.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps({k: report[k] for k in ("source", "page_count", "total_chars", "likely_scanned")},
                     ensure_ascii=False, indent=2))
    print(f"wrote {out / 'inventory.json'}")


def extract(args) -> None:
    fitz = require_fitz()
    pdf = Path(args.pdf).expanduser().resolve()
    if not pdf.is_file():
        raise SystemExit(f"PDF not found: {pdf}")
    out = Path(args.out).expanduser().resolve()
    imgs_dir = out / "images"
    pages_dir = out / "pages"
    imgs_dir.mkdir(parents=True, exist_ok=True)
    pages_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(str(pdf))
    if doc.needs_pass:
        raise SystemExit("PDF is password-protected; decrypt it first.")
    selected = parse_pages(args.pages, doc.page_count)
    data = {
        "source": str(pdf),
        "page_count": doc.page_count,
        "selected_pages": selected,
        "dpi_for_page_previews": args.dpi,
        "pages": [],
    }

    for n in selected:
        page = doc[n - 1]
        item = page_dict(page, fitz)
        seen: set[int] = set()
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                info = doc.extract_image(xref)
            except Exception as exc:
                item["images"].append({"xref": xref, "error": str(exc)})
                continue
            ext = info.get("ext", "png")
            name = f"p{n:02d}_x{xref}.{ext}"
            (imgs_dir / name).write_bytes(info["image"])
            try:
                rects = [list(map(lambda v: round(v, 2), r)) for r in page.get_image_rects(xref)]
            except Exception:
                rects = []
            item["images"].append(
                {
                    "xref": xref,
                    "file": f"images/{name}",
                    "width_px": info["width"],
                    "height_px": info["height"],
                    "ext": ext,
                    "rects": rects,
                }
            )
        pix = page.get_pixmap(dpi=args.dpi, colorspace=fitz.csRGB)
        preview = f"pages/page-{n:02d}.png"
        pix.save(str(out / preview))
        item["preview"] = preview
        data["pages"].append(item)
        print(f"page {n}/{doc.page_count}: {len(item['text'])} chars, {len(item['images'])} image(s)")

    (out / "document.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"wrote {out / 'document.json'}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_inv = sub.add_parser("inventory", help="count pages/chars/images without writing page images")
    p_inv.add_argument("--pdf", required=True)
    p_inv.add_argument("--out", required=True)
    p_inv.set_defaults(func=inventory)

    p_ext = sub.add_parser("extract", help="extract text, blocks, images, and page previews")
    p_ext.add_argument("--pdf", required=True)
    p_ext.add_argument("--out", required=True)
    p_ext.add_argument("--pages", default="all", help="e.g. all or 1-3,5")
    p_ext.add_argument("--dpi", type=int, default=200, help="page preview DPI (default 200)")
    p_ext.set_defaults(func=extract)
    return parser


def main() -> None:
    args = build_parser().parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
