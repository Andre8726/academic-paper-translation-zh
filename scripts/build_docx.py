#!/usr/bin/env python3
"""Build a Chinese academic translation DOCX from a JSON spec.

Requires python-docx. Run with the DSH runtime Python (or any Python with
python-docx installed). Paths in the spec are resolved relative to the spec file.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ACCENT = RGBColor(0xA6, 0x19, 0x2E)
GRAY = RGBColor(0x59, 0x59, 0x59)
DARK = RGBColor(0x22, 0x22, 0x22)
BLACK = RGBColor(0, 0, 0)
TOKEN_RE = re.compile(r"(~[^~]+~|\^[^\^]+\^)")
COMBINED_RE = re.compile(r"([A-Za-z\u0370-\u03ff\u0394])~([^~]+)~\^([^\^]+)\^")
ACCENT_RE = re.compile(r"([A-Za-z\u0370-\u03ff\u0394])([\u0302\u20d7\u0303])")
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


class Builder:
    def __init__(self, spec: dict, spec_dir: Path):
        self.spec = spec
        self.base = spec_dir
        self.f = {
            "body_east": "Songti SC",
            "head_east": "Heiti SC",
            "ascii": "Times New Roman",
        }
        self.f.update(spec.get("fonts", {}))
        self.doc = Document()
        self._setup_page()
        self._setup_styles()
        self._setup_header_footer()

    # ---------- text ----------
    @staticmethod
    def _m_run(text):
        from xml.sax.saxutils import escape
        return f'<m:r><m:t>{escape(str(text))}</m:t></m:r>'

    def _m_base(self, base, accent=None):
        if accent:
            from xml.sax.saxutils import escape
            return (f'<m:acc><m:accPr><m:chr m:val="{accent}"/></m:accPr>'
                    f'<m:e>{self._m_run(base)}</m:e></m:acc>')
        return self._m_run(base)

    def add_inline_math(self, p, base, sub=None, sup=None, accent=None):
        xml = f'<m:oMath xmlns:m="{M_NS}">'
        if sub is not None and sup is not None:
            xml += (f'<m:sSubSup><m:e>{self._m_base(base, accent)}</m:e>'
                    f'<m:sub>{self._m_run(sub)}</m:sub><m:sup>{self._m_run(sup)}</m:sup></m:sSubSup>')
        elif sub is not None:
            xml += (f'<m:sSub><m:e>{self._m_base(base, accent)}</m:e>'
                    f'<m:sub>{self._m_run(sub)}</m:sub></m:sSub>')
        else:
            xml += self._m_base(base, accent)
        xml += '</m:oMath>'
        p._p.append(parse_xml(xml))

    def _add_plain(self, p, text, size, bold, color, east, ascii_font, italic):
        for tok in TOKEN_RE.split(str(text)):
            if not tok:
                continue
            sub = sup = False
            if tok.startswith("~") and tok.endswith("~") and len(tok) > 2:
                sub = True; tok = tok[1:-1]
            elif tok.startswith("^") and tok.endswith("^") and len(tok) > 2:
                sup = True; tok = tok[1:-1]
            r = p.add_run(tok)
            tok_font = "Songti SC" if re.search(r"[\u0370-\u03ff\u0300-\u036f\u20d0-\u20ff]", tok) else ascii_font
            self._font(r, east=east, ascii_font=tok_font, size=size,
                       bold=bold, italic=italic, color=color)
            if sub: r.font.subscript = True
            if sup: r.font.superscript = True

    def _add_segment(self, p, text, size, bold, color, east, ascii_font, italic):
        text = str(text)
        pos = 0
        for m in ACCENT_RE.finditer(text):
            self._add_plain(p, text[pos:m.start()], size, bold, color, east, ascii_font, italic)
            mark = {"\u20d7": "→", "\u0302": "^", "\u0303": "~"}.get(m.group(2), m.group(2))
            self.add_inline_math(p, m.group(1), accent=mark)
            pos = m.end()
        self._add_plain(p, text[pos:], size, bold, color, east, ascii_font, italic)

    def add_runs(self, p, text, size=10.5, bold=False, color=BLACK,
                 east=None, ascii_font=None, italic=False):
        east = east or self.f["body_east"]
        ascii_font = ascii_font or self.f["ascii"]
        text = str(text)
        pos = 0
        # Base~sub~^sup^ must be one OMML object; separate runs put sup far right.
        for m in COMBINED_RE.finditer(text):
            self._add_segment(p, text[pos:m.start()], size, bold, color, east, ascii_font, italic)
            base, sub, sup = m.group(1), m.group(2), m.group(3)
            self.add_inline_math(p, base, sub=sub, sup=sup)
            pos = m.end()
        tail = text[pos:]
        self._add_segment(p, tail, size, bold, color, east, ascii_font, italic)

    @staticmethod
    def _font(run, east, ascii_font, size=None, bold=None, italic=None, color=None):
        run.font.name = ascii_font
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:ascii"), ascii_font)
        rFonts.set(qn("w:hAnsi"), ascii_font)
        rFonts.set(qn("w:eastAsia"), east)
        if size is not None:
            run.font.size = Pt(size)
        if bold is not None:
            run.font.bold = bold
        if italic is not None:
            run.font.italic = italic
        if color is not None:
            run.font.color.rgb = color

    @staticmethod
    def _para(p, before=0, after=4, line=1.5, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              first_indent=None, left=0, right=0, keep_next=False, keep_together=False):
        pf = p.paragraph_format
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.line_spacing = line
        pf.alignment = align
        pf.left_indent = Cm(left)
        pf.right_indent = Cm(right)
        pf.keep_with_next = keep_next
        pf.keep_together = keep_together
        if first_indent is not None:
            pf.first_line_indent = Cm(first_indent)

    # ---------- document setup ----------
    def _setup_page(self):
        cfg = self.spec.get("page", {})
        sec = self.doc.sections[0]
        sec.page_width = Cm(cfg.get("width_cm", 21.0))
        sec.page_height = Cm(cfg.get("height_cm", 29.7))
        m = cfg.get("margins_cm", [2.15, 2.0, 2.35, 2.35])
        sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = map(Cm, m)
        sec.header_distance = Cm(0.9)
        sec.footer_distance = Cm(0.9)

    def _setup_styles(self):
        normal = self.doc.styles["Normal"]
        normal.font.name = self.f["ascii"]
        normal.font.size = Pt(self.spec.get("body_size", 10.5))
        normal.font.color.rgb = BLACK
        rPr = normal.element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), self.f["body_east"])
        rFonts.set(qn("w:ascii"), self.f["ascii"])
        rFonts.set(qn("w:hAnsi"), self.f["ascii"])

        for lvl in (1, 2):
            st = self.doc.styles[f"Heading {lvl}"]
            st.font.name = "Arial"
            st.font.bold = True
            st.font.color.rgb = BLACK
            rPr = st.element.get_or_add_rPr()
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                rPr.append(rFonts)
            rFonts.set(qn("w:eastAsia"), self.f["head_east"])
            rFonts.set(qn("w:ascii"), "Arial")
            rFonts.set(qn("w:hAnsi"), "Arial")

    def _add_field(self, p, instr):
        run = p.add_run()
        f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
        f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
        t = OxmlElement("w:t"); t.text = "1"
        f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
        for el in (f1, it, f2, t, f3):
            run._r.append(el)
        self._font(run, self.f["body_east"], self.f["ascii"], size=8, color=GRAY)

    def _setup_header_footer(self):
        sec = self.doc.sections[0]
        header_text = self.spec.get("header") or (
            "中文译稿  ·  " + str(self.spec.get("title", "")).replace("\n", " ")
        )
        hp = sec.header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        hp.paragraph_format.space_after = Pt(2)
        self.add_runs(hp, header_text, size=7.5, color=GRAY)
        pPr = hp._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "4")
        bottom.set(qn("w:space"), "2"); bottom.set(qn("w:color"), "BFBFBF")
        pBdr.append(bottom); pPr.append(pBdr)

        fp = sec.footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.add_runs(fp, "第 ", size=8, color=GRAY)
        self._add_field(fp, " PAGE ")
        self.add_runs(fp, " 页 / 共 ", size=8, color=GRAY)
        self._add_field(fp, " NUMPAGES ")
        self.add_runs(fp, " 页", size=8, color=GRAY)

    # ---------- front matter ----------
    def title_block(self):
        s = self.spec
        p = self.doc.add_paragraph()
        self._para(p, before=0, after=7, line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
        self.add_runs(p, s["title"], size=s.get("title_size", 19.5), bold=True,
                      east=self.f["head_east"], ascii_font="Arial")

        if s.get("original_title"):
            p = self.doc.add_paragraph()
            self._para(p, before=0, after=5, line=1.12, align=WD_ALIGN_PARAGRAPH.CENTER)
            self.add_runs(p, "原文：" + s["original_title"], size=8.6, italic=True, color=GRAY)

        if s.get("authors"):
            p = self.doc.add_paragraph()
            self._para(p, before=0, after=4, line=1.2, align=WD_ALIGN_PARAGRAPH.CENTER)
            self.add_runs(p, s["authors"], size=10.2)

        for a in s.get("affiliations", []):
            p = self.doc.add_paragraph()
            self._para(p, before=0, after=1.5, line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
            self.add_runs(p, a, size=8.3, color=RGBColor(0x33, 0x33, 0x33))

        for text, color, size in (
            (s.get("correspondence", ""), RGBColor(0x33, 0x33, 0x33), 8.3),
            (s.get("doi", ""), ACCENT, 8.6),
        ):
            if text:
                p = self.doc.add_paragraph()
                self._para(p, before=4 if text == s.get("doi") else 0, after=4,
                           line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
                self.add_runs(p, text, size=size, color=color)

        if s.get("note"):
            p = self.doc.add_paragraph()
            self._para(p, before=1, after=4, line=1.05, align=WD_ALIGN_PARAGRAPH.CENTER)
            self.add_runs(p, s["note"], size=7.6, italic=True, color=GRAY)

    def label(self, text):
        p = self.doc.add_paragraph()
        self._para(p, before=6, after=2, line=1.0, align=WD_ALIGN_PARAGRAPH.LEFT, keep_next=True)
        self.add_runs(p, text, size=10, bold=True, east=self.f["head_east"],
                      ascii_font="Arial", color=ACCENT)

    def abstract(self):
        if not self.spec.get("abstract"):
            return
        self.label("摘要")
        p = self.doc.add_paragraph()
        self._para(p, before=4, after=8, line=1.42, first_indent=0, left=0.25, right=0.25)
        self.add_runs(p, self.spec["abstract"], size=9.8)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), "F4F4F4")
        pPr.append(shd)
        pBdr = OxmlElement("w:pBdr")
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single"); left.set(qn("w:sz"), "18")
        left.set(qn("w:space"), "10"); left.set(qn("w:color"), "A6192E")
        pBdr.append(left); pPr.append(pBdr)

        if self.spec.get("keywords"):
            p = self.doc.add_paragraph()
            self._para(p, before=0, after=5, line=1.3, align=WD_ALIGN_PARAGRAPH.LEFT)
            self.add_runs(p, "关键词：", size=9.5, bold=True, east=self.f["head_east"])
            self.add_runs(p, self.spec["keywords"], size=9.5)

    # ---------- content ----------
    def heading(self, text, level=1):
        p = self.doc.add_paragraph(style=f"Heading {level}")
        pf = p.paragraph_format
        if level == 1:
            pf.space_before = Pt(14); pf.space_after = Pt(6); pf.line_spacing = 1.2
            size = 15
        else:
            pf.space_before = Pt(10); pf.space_after = Pt(4); pf.line_spacing = 1.2
            size = 11.5
        pf.keep_with_next = True
        pf.outline_level = 0 if level == 1 else 1
        self.add_runs(p, text, size=size, bold=True, east=self.f["head_east"],
                      ascii_font="Arial", color=ACCENT if level == 1 else BLACK)

    def body(self, text, indent=True):
        p = self.doc.add_paragraph()
        self._para(p, before=0, after=5, line=1.48,
                   first_indent=0.74 if indent else 0)
        self.add_runs(p, text, size=self.spec.get("body_size", 10.5))

    def path(self, value):
        p = Path(value).expanduser()
        return p if p.is_absolute() else (self.base / p)

    def picture(self, value, width_cm, before=8, after=1, keep_next=True):
        p = self.doc.add_paragraph()
        self._para(p, before=before, after=after, line=1.0,
                   align=WD_ALIGN_PARAGRAPH.CENTER, keep_next=keep_next, keep_together=True)
        p.add_run().add_picture(str(self.path(value)), width=Cm(width_cm))

    def caption(self, lead, rest, size=9.0, after=8):
        p = self.doc.add_paragraph()
        self._para(p, before=1, after=after, line=1.15, align=WD_ALIGN_PARAGRAPH.JUSTIFY)
        r = p.add_run(lead)
        self._font(r, east=self.f["body_east"], ascii_font=self.f["ascii"],
                   size=size, bold=True, color=BLACK)
        self.add_runs(p, " " + rest, size=size, color=DARK)

    @staticmethod
    def _shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill)
        tcPr.append(shd)

    @staticmethod
    def _borders(table, color="A6A6A6", sz="4"):
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "single"); el.set(qn("w:sz"), sz)
            el.set(qn("w:space"), "0"); el.set(qn("w:color"), color)
            borders.append(el)
        table._tbl.tblPr.append(borders)

    @staticmethod
    def _cell_margins(table, top=55, bottom=55, left=90, right=90):
        mar = OxmlElement("w:tblCellMar")
        for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
            el = OxmlElement(f"w:{side}")
            el.set(qn("w:w"), str(val)); el.set(qn("w:type"), "dxa")
            mar.append(el)
        table._tbl.tblPr.append(mar)

    def table(self, block):
        lead = block.get("lead")
        if lead:
            p = self.doc.add_paragraph()
            self._para(p, before=2, after=4, line=1.1,
                       align=WD_ALIGN_PARAGRAPH.CENTER, keep_next=True)
            self.add_runs(p, lead, size=9.6, bold=True, east=self.f["head_east"])

        headers = block.get("headers", [])
        rows = block.get("rows", [])
        if not headers:
            raise ValueError("table block needs headers")
        table = self.doc.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = block.get("widths_cm") or [16.0 / len(headers)] * len(headers)
        if len(widths) != len(headers):
            raise ValueError("widths_cm length must match headers")
        widths_cm = [Cm(w) for w in widths]
        header_size = block.get("header_size", 8.9)
        body_size = block.get("body_size", 8.4)

        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.width = widths_cm[i]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(1)
            self.add_runs(p, h, size=header_size, bold=True)
            self._shade(cell, block.get("header_fill", "E8EEF4"))

        for row in rows:
            if len(row) != len(headers):
                raise ValueError(f"table row has {len(row)} cells, expected {len(headers)}")
            cells = table.add_row().cells
            for i, val in enumerate(row):
                cells[i].width = widths_cm[i]
                p = cells[i].paragraphs[0]
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if i in block.get("center_cols", [1, 2])
                               else WD_ALIGN_PARAGRAPH.LEFT)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.05
                self.add_runs(p, val, size=body_size)

        self._borders(table, color=block.get("border_color", "A6A6A6"))
        self._cell_margins(table)
        for row in table.rows:
            trPr = row._tr.get_or_add_trPr()
            cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true")
            trPr.append(cant)
        tr = table.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader"); th.set(qn("w:val"), "true")
        trPr.append(th)
        if block.get("note"):
            p = self.doc.add_paragraph()
            self._para(p, before=4, after=6, line=1.15, align=WD_ALIGN_PARAGRAPH.LEFT)
            self.add_runs(p, block["note"], size=8.3, color=RGBColor(0x33, 0x33, 0x33))

    def equation_omml(self, block):
        """Insert native Office Math (OMML) rows in a borderless 2-column table."""
        files = block.get("files") or []
        if not files:
            raise ValueError("equation_omml needs files")
        table = self.doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = [Cm(block.get("formula_width_cm", 15.2)), Cm(block.get("number_width_cm", 1.2))]
        for i, w in enumerate(widths):
            table.columns[i].width = w
            table.rows[0].cells[i].width = w
        # no borders, tight margins
        tblPr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}"); el.set(qn("w:val"), "nil"); borders.append(el)
        tblPr.append(borders)
        mar = OxmlElement("w:tblCellMar")
        for side, val in (("top", 60), ("left", 20), ("bottom", 60), ("right", 20)):
            el = OxmlElement(f"w:{side}"); el.set(qn("w:w"), str(val)); el.set(qn("w:type"), "dxa"); mar.append(el)
        tblPr.append(mar)
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)

        left, right = table.rows[0].cells
        left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = left.paragraphs[0]
        self._para(p, before=0, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        for idx, f in enumerate(files):
            if idx:
                p = left.add_paragraph()
                self._para(p, before=0, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
            xml = Path(f).read_text(encoding="utf-8")
            p._p.append(parse_xml(xml))
        p = right.paragraphs[0]
        self._para(p, before=0, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
        self.add_runs(p, f"({block.get('number', '')})", size=block.get("number_size", 10.5))
        trPr = table.rows[0]._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit"); cant.set(qn("w:val"), "true"); trPr.append(cant)

    def refs(self, block):
        if block.get("page_break_before"):
            self.doc.add_page_break()
        heading = block.get("heading", "参考文献")
        if heading:
            self.heading(heading, 1)
        size = block.get("font_size", 7.9)
        numbering = block.get("numbering", "paren")
        for i, item in enumerate(block.get("items", []), 1):
            if isinstance(item, dict):
                n = item.get("n", i); text = item.get("text", "")
            else:
                n = i; text = str(item)
            p = self.doc.add_paragraph()
            self._para(p, before=0, after=2.2, line=1.06,
                       align=WD_ALIGN_PARAGRAPH.LEFT, left=0.62)
            p.paragraph_format.first_line_indent = Cm(-0.62)
            if numbering == "paren":
                prefix = f"({n}) "
            elif numbering == "bracket":
                prefix = f"[{n}] "
            else:
                prefix = ""
            self.add_runs(p, prefix, size=size)
            self.add_runs(p, text, size=size)

    def toc_figure(self):
        cfg = self.spec.get("toc_figure")
        if not cfg:
            return
        self.picture(cfg["path"], cfg.get("width_cm", 12.6), before=4, after=1, keep_next=True)
        self.caption(cfg.get("lead", "图文摘要."), cfg.get("caption", ""), size=8.0, after=2)

    def build(self):
        self.title_block()
        self.abstract()
        self.toc_figure()
        page_break_after_front = self.spec.get("page_break_after_front", True)
        if page_break_after_front and self.spec.get("blocks"):
            # A front-matter page break is only inserted if the first block is a body heading.
            first = self.spec["blocks"][0]
            if first.get("type") == "h1":
                self.doc.add_page_break()

        for block in self.spec.get("blocks", []):
            kind = block.get("type")
            if kind == "h1":
                self.heading(block["text"], 1)
            elif kind == "h2":
                self.heading(block["text"], 2)
            elif kind == "p":
                self.body(block["text"], indent=block.get("indent", True))
            elif kind in ("figure", "image"):
                self.picture(block["path"], block.get("width_cm", 14.0),
                             before=block.get("before", 8), after=block.get("after", 1),
                             keep_next=True)
                if block.get("caption"):
                    self.caption(block.get("lead", ""), block["caption"],
                                 size=block.get("caption_size", 9.0))
            elif kind == "equation":
                self.picture(block["path"], block.get("width_cm", 15.0),
                             before=block.get("before", 2), after=block.get("after", 2),
                             keep_next=True)
            elif kind == "equation_omml":
                self.equation_omml(block)
            elif kind == "table":
                self.table(block)
            elif kind == "pagebreak":
                self.doc.add_page_break()
            elif kind == "refs":
                self.refs(block)
            else:
                raise ValueError(f"unknown block type: {kind}")

        self.doc.core_properties.title = self.spec.get("title", "")
        self.doc.core_properties.subject = self.spec.get("subject", "中文翻译")
        return self.doc


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--spec", required=True, help="translation spec JSON")
    ap.add_argument("--out", required=True, help="output .docx")
    args = ap.parse_args()

    spec_path = Path(args.spec).expanduser().resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    doc = Builder(spec, spec_path.parent).build()
    out = Path(args.out).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print(out)


if __name__ == "__main__":
    main()
