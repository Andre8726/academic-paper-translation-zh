# academic-paper-translation-zh

把英文学术论文或补充信息 PDF **完整翻译为可检索的简体中文 A4 PDF**，保留原始插图、图注、公式、表格和参考文献。先重建为 Word，再导出 PDF；不是图片型翻译。

## 适合用它做什么

- 翻译一篇英文论文，输出排版整洁的中文 PDF。
- 正文与补充信息可分别输出，补充图、补充表不遗漏。
- 展示公式优先重建为原生 Word 公式（OMML，可选中、可编辑）；表格重建为原生 Word 表格。
- 图内英文刻度/图例保持原样，图注和表注翻译为中文。
- 参考文献保留英文原文，便于核对。

## 不适合

- 只翻译摘要或部分段落。
- 扫描版 PDF 的 OCR；需要先做 OCR 再进入本流程。
- 逐字审校、正式同行评审或图片型生图翻译。

## 输出格式

- A4 纵向；正文宋体 10.5 pt，标题黑体，强调色深红 `#A6192E`。
- 首页：中文标题、英文原题、作者、单位、通讯作者、DOI、摘要、关键词、图文摘要。
- 正文：分节标题、首行缩进、图/图注、公式、原生表格、页眉页脚。
- 参考文献：保留英文，悬挂缩进，编号连续。
- 最终交付 PDF；DOCX 仅作为可编辑源文件。

## 安装

作为 DSH 用户级 skill：

```bash
mkdir -p ~/.dsh/skills
git clone https://github.com/Andre8726/academic-paper-translation-zh.git \
  ~/.dsh/skills/academic-paper-translation-zh
```

如果已经下载本目录，也可以直接复制：

```bash
cp -R academic-paper-translation-zh ~/.dsh/skills/
```

安装后在会话中直接说：

```text
用 academic-paper-translation-zh 翻译这两个 PDF，保留图片、公式和表格。
```

或使用 `/academic-paper-translation-zh`。

## 依赖

- Python 3.9+
- PyMuPDF（`fitz`）
- python-docx
- matplotlib（公式渲染）
- Pillow（页面检查图）
- LibreOffice Kit（DSH 的 `office-export-pdf` / `office-render` skill 提供）

macOS 上如果 DSH 运行时 Python 加载 PyMuPDF 报代码签名错误，使用：

```bash
./scripts/bootstrap_pymupdf.sh "$PWD/work/translation"
```

## 使用流程

1. `extract_pdf.py inventory`：统计页数、字符数、图片数，判断是否为扫描件。
2. `extract_pdf.py extract`：导出 `document.json`、页面预览和内嵌图。
3. 逐页翻译，写成 `translation.json`（见 `references/spec-example.json`）。
4. `latex_to_omml.py`：把展示公式转成 Word 原生 OMML；仅在转换失败时用 `render_equation.py` 生成 PNG。
5. `build_docx.py`：生成中文 DOCX。
6. LibreOffice Kit 导出 PDF。
7. `inspect_pdf.py`：渲染全部页面，检查图、表、公式、字体和空白页。
8. 交付 PDF。

详细规则见 [SKILL.md](SKILL.md) 和 [translation-style.md](references/translation-style.md)。

## 目录

```text
SKILL.md
references/
  spec-example.json
  translation-style.md
scripts/
  bootstrap_pymupdf.sh
  build_docx.py
  extract_pdf.py
  inspect_pdf.py
  latex_to_omml.py
  render_equation.py
agents/openai.yaml
```

## 说明

- 本 skill 只处理用户有权使用的文档。
- 翻译稿是参考译稿，专业术语与数据应以英文原文为准。
- 图内英文标注刻意保留，避免重绘造成数据失真。
