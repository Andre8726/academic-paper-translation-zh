---
name: academic-paper-translation-zh
description: 将英文学术论文或补充信息 PDF 完整翻译为简体中文 A4 排版 PDF，保留原始插图、图注、公式、表格和参考文献；先重建为 Word 再导出，正文可检索，不是图片型 PDF。适用于“翻译论文 PDF”“补充材料翻译”“中文译稿保留图片”“按这种排版做译稿”。
whenToUse: 当用户给出英文论文/补充信息 PDF，要求中文 PDF、保留图表、公式或参考文献时使用；不用于只翻译摘要、只做图片型生图翻译或逐字审校。
---

# 学术论文 PDF 中文排版翻译

把英文论文重建为**可检索的中文 A4 PDF**：作者/单位、摘要、关键词、正文、图、图注、公式、表格、作者信息、致谢和参考文献都保留。不是把英文页面直接盖中文，也不是图片型 PDF；插图沿用原图，图注与表注译为中文。

本 skill 不绑定学科、期刊或某篇论文：术语表、章节结构和缩写均按当前论文建立，下面出现的字段只是排版示例。

## 交付契约

- 一篇论文一个 PDF；正文与补充信息通常分别输出。只有用户明确要求合并时才合并。
- A4 纵向，页边距约 2.15/2.0/2.35/2.35 cm；正文宋体 10.5 pt，标题黑体，强调色深红 `#A6192E`。
- 图必须真实出现在译稿中：优先提取 PDF 内嵌图；矢量图按图区 300 dpi 渲染。不要用正文截图代替插图，不要重绘数据图。
- 图内英文坐标、刻度、图例保持原样；图注、表注、正文全部中文。首页说明“图内英文标注未改动”。
- 公式优先用 matplotlib mathtext 渲染为透明/白底 PNG，再插入 Word；简单行内符号用 `~下标~`、`^上标^` 标记。
- 表格重建为 Word 原生表格，不截图；表头跨页重复，行不跨页拆分。
- 参考文献保留英文原文，不翻译；作者单位、通讯作者、致谢、利益冲突声明翻译为中文。
- 最终只交付 PDF（可另留 DOCX 作源文件，但不要主动作为交付物）。

## 变量约定

```bash
SKILL="$HOME/.dsh/skills/academic-paper-translation-zh"
JOB="$PWD/work/translation"      # 本次任务目录
PDF="/path/to/paper.pdf"
OUT="$PWD/output/paper_zh.pdf"
```

`$DSH_PYTHON` 由 `load_workspace_dependencies` 返回；`$NODE` 与 `$LIBREOFFICE_CLI` 来自 `office-export-pdf` / `office-render`。

## 环境与依赖

- PDF 解析：PyMuPDF（`fitz`）。macOS 上如果 DSH 运行时 Python 报 `_extra.so` 代码签名错误，改用 `/usr/bin/python3`，并把 PyMuPDF 安装到任务目录：
  ```bash
  /usr/bin/python3 -m pip install --target "$JOB/.pymupdf" "pymupdf==1.24.14"
  export PYTHONPATH="$JOB/.pymupdf"
  ```
  也见 `scripts/bootstrap_pymupdf.sh`。不要在未说明的情况下更换渲染器。
- DOCX：DSH 运行时 Python 的 `python-docx`。先调用 `load_workspace_dependencies` 获取 Python 绝对路径，或使用当前 profile 的 `.../dependencies/python/bin/python3`。
- 公式：`/usr/bin/python3` + matplotlib（`MPLCONFIGDIR` 指向任务目录，避免写 `~/.matplotlib`）。
- PDF 导出与渲染：遵循 `office-export-pdf` / `office-render` skill 的 LibreOffice Kit 路径。不要用系统 `soffice` 替代。
- 字体：macOS `Songti SC`（正文）、`Heiti SC`（标题）、`Times New Roman`（西文）。导出时确认 `missingFonts` 为空。

## 工作流

### 1. 盘点输入

先确认文件、页数、是否有补充信息、输出目录。对每个 PDF 统计每页字符数与图片数：

```bash
PYTHONPATH="$JOB/.pymupdf" /usr/bin/python3 "$SKILL/scripts/extract_pdf.py" inventory \
  --pdf "$PDF" --out "$JOB/inventory"
```

- 文本量正常（每页通常 >500 字符）＝ born-digital，直接抽取。
- 整页几乎没有文本、主要是扫描图＝先告知需要 OCR；不要凭空翻译。
- 记录每页图片 xref、bbox、像素尺寸，用于判断图是内嵌位图还是矢量绘制。

### 2. 抽取文本与图片

```bash
PYTHONPATH="$JOB/.pymupdf" /usr/bin/python3 "$SKILL/scripts/extract_pdf.py" extract \
  --pdf "$PDF" --out "$JOB/extract" --dpi 200
```

生成：

- `document.json`：逐页 `text`、文本块 bbox、图片清单、页面尺寸。
- `pages/page-01.png`：整页预览，仅用于核对版式，不直接放进译稿。
- `images/p01_x102.png`：内嵌图原图；`xref` 与 bbox 写入 JSON。

逐页读完 `document.json` 的正文，不要只读第一页或摘要。遇到公式乱码时，回到整页预览读取公式结构。

### 3. 选图

- 有内嵌图：直接用 `images/` 中的原图，分辨率通常足够。
- 矢量图/多面板图：按整页预览确定图区，用 PyMuPDF `clip` 以 300 dpi 渲染，不要包含正文与图注。
- 补充信息常见“一页一张图”：按图区裁切；图内英文保留。
- 记录每张图对应的图号，确保正文图号、图注、图片三者一致。

### 4. 完整翻译

按原文顺序逐段翻译，不删段、不合并结果、不写摘要替代正文。规则见 [translation-style.md](references/translation-style.md)。

要点：

- 数字、单位、误差、样品/模型编号、方向与坐标、图号、表号必须与原文一致。
- 首次出现的缩写写“中文（英文缩写）”，后文用缩写。
- 先建立本文术语表：同一术语全文一种译法；首次出现给英文原词；缩写首次出现给全称。不要把某一学科的译名硬套到其他领域。
- 公式与变量用 `~...~` / `^...^` 标记，例如 `x~i~`、`σ~max~`、`E~a~^eff^`。
- 参考文献不翻译；只清理 PDF 断行造成的 `well- known`、`1234− 1240` 等伪空格。
- 图注翻译为“图 1. ……”或“图 S1. ……”，表注为“表 1 ……”或“表 S1 ……”，句末用中文句号。

### 5. 公式：优先原生 OMML

目标是把展示公式写成 **Word 原生公式（OMML）**，在 PDF 中可选中、可编辑、矢量显示；不要把公式截图或 PNG 作为默认方案。

先把公式写成 JSON：

```json
{
  "eq1": {"number": "1", "rows": ["\\frac{d\\hat{m}}{dt}=\\gamma\\hat{m}\\times\\frac{dF}{d\\hat{m}}"]},
  "eq2": {"number": "2", "rows": ["E=mc^2"]}
}
```

`rows` 中每一行是一条 LaTeX；多行公式就写多条。转换：

```bash
"$DSH_PYTHON" -m pip install --target "$JOB/.mathdeps" latex2mathml
PYTHONPATH="$JOB/.mathdeps" "$DSH_PYTHON" "$SKILL/scripts/latex_to_omml.py" \
  --input "$JOB/equations.json" --out "$JOB/eq_omml"
```

输出 `eq_omml/<name>_<i>.xml` 与 `eq_omml/omml_map.json`。spec 中使用：

```json
{"type": "equation_omml", "number": "1", "files": ["eq_omml/eq1_0.xml"]}
```

- 公式编号放在右侧；多行公式共用一个编号。
- 转换依赖 Microsoft Word 的 `mathml2omml.xsl`；macOS 默认路径已内置，其他平台用 `--xsl` 指定。
- 只有当 LaTeX 过于复杂、OMML 转换失败时，才退回 `render_equation.py` 生成 PNG，并在 spec 中使用 `{"type": "equation", "path": ...}`。
- 不要用整页截图代替公式，也不要把公式放进正文图片。

### 6. 写翻译 spec

把译文写成 JSON（结构见 [spec-example.json](references/spec-example.json)），不要把长文直接塞进 shell 参数。路径相对 spec 文件解析。

最小结构：

```json
{
  "title": "中文标题",
  "original_title": "English title",
  "authors": "Author A¹, Author B²*",
  "affiliations": ["¹ 单位，城市 邮编", "² 单位，城市 邮编"],
  "correspondence": "* 通讯作者：a@example.com",
  "doi": "Journal · https://doi.org/...",
  "note": "说明：中文参考译稿；插图沿用原文，图内英文标注未作改动。",
  "abstract": "摘要正文……",
  "keywords": "关键词1；关键词2",
  "header": "中文译稿 · 中文短标题",
  "toc_figure": {"path": "figs/toc.png", "width_cm": 12.6, "lead": "图文摘要.", "caption": "……"},
  "blocks": [
    {"type": "h1", "text": "引言"},
    {"type": "p", "text": "正文……"},
    {"type": "figure", "path": "figs/fig1.png", "width_cm": 16.2, "lead": "图 1.", "caption": "……"},
    {"type": "equation_omml", "number": "1", "files": ["eq_omml/eq1_0.xml"]},
    {"type": "table", "lead": "表 S1  ……", "headers": ["项目", "数值"], "rows": [["...", "..."]], "widths_cm": [8.0, 8.0], "note": "注：……"},
    {"type": "pagebreak"},
    {"type": "refs", "page_break_before": true, "items": ["(1) ...", "(2) ..."]}
  ]
}
```

块类型：`h1`、`h2`、`p`、`figure`、`equation`、`table`、`pagebreak`、`refs`。`p` 默认首行缩进 2 字符；`indent: false` 可关闭。

### 7. 生成 DOCX

```bash
"$DSH_PYTHON" "$SKILL/scripts/build_docx.py" \
  --spec "$JOB/translation.json" --out "$JOB/translation.docx"
```

生成后不要急着交付；先做结构检查：标题/作者/摘要/关键词是否存在，图号顺序是否连续，表格是否完整，参考文献条数是否与原文一致。

### 8. 导出 PDF

按 `office-export-pdf` skill 的路径执行：

```bash
"$NODE" "$LIBREOFFICE_CLI" convert --input "$JOB/translation.docx" --output "$OUT.pdf"
```

检查返回 JSON 的 `missingFonts` 为空、输出文件以 `%PDF-` 开头且非空。不要重复转换未修改的源文件。

### 9. 渲染与验收

按 `office-render` 的约定，用任务内临时目录渲染全部页面：

```bash
PYTHONPATH="$JOB/.pymupdf" /usr/bin/python3 "$SKILL/scripts/inspect_pdf.py" \
  --pdf "$OUT.pdf" --out "$JOB/render" --dpi 110
```

逐页查看 `contact-*.png` 或 `page-*.png`，检查：

- 每页无溢出、无重叠、无孤立标题；图与图注同页。
- 表格没有被分页切断；表头重复正常。
- 公式没有超出正文宽度；上下标没有丢字。
- 所有图都出现，图号与正文引用一致。
- 中文没有方框/乱码；字体已嵌入。
- 页脚总页数正确；没有意外的空白页或重复章节。

发现问题→改 spec/图→重建 DOCX→重新导出→重新渲染。不要只改 PDF 文本层。

### 10. 交付

最终回复只给 1–2 个 PDF 链接，简要说明页数、包含的图/表。不要把预览图、检查报告、JSON、DOCX 一并作为交付物；除非用户明确要源文件。

## 质量门槛

- 正文覆盖率：原文每个一级段落/小节在译稿中有对应内容；不得省略方法、数据可用性、作者信息、致谢等原文已有的部分。
- 图覆盖率：原文所有编号图（含补充图）都出现；图注完整；不得用一张拼图替代多个图号。
- 表覆盖率：所有编号表重建；数值、单位、上下标、脚注正确。
- 公式覆盖率：论文中的展示公式、编号公式和变量定义保留；优先使用可编辑的 OMML，上下标与符号一致。
- 参考文献：条数一致，编号连续，保留英文；不把引用编号误当正文。
- 版式：无表格跨页、无图注分离、无标题孤行、无文字溢出页面。
- 字体：`missingFonts` 为空；PDF 文本可搜索。

## 资源

- [翻译与术语规范](references/translation-style.md)
- [JSON spec 示例](references/spec-example.json)
- [PDF 抽取脚本](scripts/extract_pdf.py)
- [DOCX 构建脚本](scripts/build_docx.py)
- [公式渲染脚本](scripts/render_equation.py)
- [PDF 验收脚本](scripts/inspect_pdf.py)
- [PyMuPDF 安装脚本](scripts/bootstrap_pymupdf.sh)
