# academic-paper-translation-zh

Translate English academic papers or supporting-information PDFs into a **searchable, typeset Simplified Chinese A4 PDF**, preserving the original figures, captions, equations, tables, and references. The document is rebuilt in Word and then exported to PDF; it is not an image-only translation.

## Example output

![English and Chinese PDF excerpts from Okubo et al. shown side by side](docs/demo/translation-preview.webp)

The English paper excerpt is on the left. The Chinese excerpt is on the right, translated, typeset, and exported as an A4 PDF using this workflow. Text in the exported PDF is searchable and selectable.

The source paper is by Ren Okubo, Yu Fujikata, Izumi Takahara, and Teruyasu Mizoguchi, published in *npj Computational Materials* (2026). [Original article and DOI](https://www.nature.com/articles/s41524-026-02384-4) · [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/)

## Features

- Full-paper translation with a clean academic layout.
- Separate main-paper and supporting-information PDFs.
- Embedded figures reused at original resolution; vector figures rendered at 300 dpi.
- Equations rebuilt as native Word OMML (editable/selectable) where possible, with PNG fallback; native Word tables; English references preserved.
- Page headers/footers, Chinese captions, and a validated visual pass.

## Install

```bash
mkdir -p ~/.dsh/skills
git clone https://github.com/Andre8726/academic-paper-translation-zh.git \
  ~/.dsh/skills/academic-paper-translation-zh
```

Then invoke:

```text
Use academic-paper-translation-zh to translate these PDFs into Chinese PDFs, keeping figures, equations, and tables.
```

## Requirements

Python 3.9+, PyMuPDF, python-docx, matplotlib, Pillow, and the DSH LibreOffice Kit (`office-export-pdf` / `office-render`).

See [SKILL.md](SKILL.md) for the full workflow and [translation-style.md](references/translation-style.md) for terminology and layout rules.
