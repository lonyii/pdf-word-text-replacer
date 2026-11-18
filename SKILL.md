---
name: pdf-word-text-replacer
description: 批量把 PDF 或 Word 文件转成可编辑 Word、按规则查找替换文字、给扫描件 PDF 叠加 OCR 隐藏文字层，再把结果另存到桌面时间戳新文件夹。当用户给出一个文件夹/一批 PDF/一批 Word 文件，要求先转 Word（针对 PDF）、再把某些文本替换成另一些文本、对扫描件做 OCR 识别文字、或另存输出到桌面时使用本技能。触发词：PDF转Word替换、批量替换Word文字、扫描件OCR识别文字、给PDF加文字层、把文件夹里的PDF/Word文字替换后另存、查找替换后输出到桌面。
agent_created: true
---

# PDF/Word 批量转换与文字替换

## 概述

处理一批混合的 PDF / Word 文件：PDF 先静默转成可编辑 Word，然后所有 Word 文档按用户给定规则做查找替换，结果统一另存到桌面上一个以日期时间戳命名的新文件夹。替换覆盖正文、页眉、页脚、表格、文本框，保留原格式。另有 **扫描件 OCR 文字层** 能力：用 Windows 内置中文 OCR 给纯扫描 PDF 叠加隐藏文字层，使其可搜索、可复制（保持原版式）。

## 何时使用

- 用户给出一个文件夹、一批 PDF 文件、或一批 Word 文件（`.pdf` / `.docx` / `.doc` 混合也可）。
- 用户要求：PDF 先转 Word，然后把某些文本替换成另一文本，再另存输出。
- 用户要求输出到桌面新文件夹（按时间戳命名）。
- 用户要求对**扫描件 PDF 做 OCR / 识别文字 / 加文字层**（用 `pdf_ocr_overlay.py`，见"扫描件 OCR"一节）。

## 前置条件

- PDF→Word 转换：本机已安装 **Adobe Acrobat Pro**（非 Reader）和 **Microsoft Word**；Python 装 `pywin32`（`pip install pywin32`）。
- 扫描件 OCR：Python 装 `pymupdf` 与 WinRT  OCR 模块（见下），依赖系统内置中文 OCR（Windows 10/11 自带，无需联网）。
- 脚本均在 `scripts/` 下，用 PowerShell 的 Python 3.12 执行（该环境已验证同时具备 `fitz` 与 `winrt`）。

## 工作流

### 第 1 步：明确替换规则

向用户确认要替换的文本对。用户可能一次给多组。整理成规则数组，每组：

```json
{"find": "原文", "replace": "新文"}
```

可选字段：`match_case`（区分大小写）、`whole_word`（全字匹配）、`wildcards`（通配符），默认 `false`。

### 第 2 步：写规则文件（关键，勿跳过）

**必须把规则写成 UTF-8 的 `.json 文件`**，再用 `--rules <文件路径>` 传给脚本。

不要在命令行里直接把中文 JSON 当参数传——PowerShell 5.1 传给 Python 时会破坏中文编码，导致 `JSONDecodeError`。

规则文件格式：

```json
[
  {"find": "甲方", "replace": "建设单位"},
  {"find": "乙方", "replace": "施工单位"}
]
```

把文件写到临时目录（如 `%TEMP%\replace_rules.json`），用 Write 工具写 UTF-8。

### 第 3 步：运行主脚本

用 PowerShell 调用（Windows 环境）：

```powershell
python "<skill_dir>\scripts\batch_replace.py" `
  --input "<路径1>" "<路径2>" ... `
  --rules "<规则json文件路径>"
```

`--input` 接受：单个文件、多个文件、文件夹（递归扫描子目录）、通配符，可混合。

可选参数：
- `--out-dir <路径>`：自定义输出目录（默认：桌面 `Word替换_YYYYMMDD_HHMMSS` 文件夹）。
- `--no-convert-pdf`：跳过 PDF 转换（只处理 Word 文件）。

脚本输出以 `|` 分隔的日志行：`RULES` / `INPUTS` / `OUTDIR` / `CONVERT_OK` / `CONVERT_FAIL` / `REPLACE_OK` / `REPLACE_FAIL` / `DONE`。

### 第 4 步：核对与交付

1. 检查日志里的 `DONE|N|输出目录`，确认处理数量与预期一致。
2. 抽查替换效果：用 Word COM 打开输出文档读 `Content.Text`，确认新文本存在、旧文本归零；页眉页脚也要查（`Sections(1).Headers(1).Range.Text`）。
3. 用 `present_files` 展示输出文件夹里的 docx。
4. 告知用户输出目录路径和处理文件数。

## 扫描件 OCR：给 PDF 叠加隐藏文字层

当用户要求对**扫描件 PDF** 做 "OCR / 识别文字 / 加文字层 / 变成可搜索"，用 `scripts/pdf_ocr_overlay.py`。

### 为什么不用 Acrobat 自带 OCR（重要，别再走弯路）

- Acrobat 的"识别文本"命令 `PaperCapture` 虽能通过 `app.execMenuItem` 触发，但在 IAC 外部上下文**不产生文字层**（已多次实测）；OCR 属受限命令，官方要求从"文件夹级受信任脚本"调用（需写 Program Files + 管理员权限 + 重启 Acrobat），且其语言选择对话框无法静默设置。
- 因此采用 **Windows 内置 OCR**（`Windows.Media.Ocr`，简体中文），配合 PyMuPDF，完全静默、可批量、无需联网、无需管理员。

### 用法

```powershell
# 单个扫描 PDF -> 同名 _ocr.pdf
python scripts\pdf_ocr_overlay.py --input "D:\scans\扫描件.pdf"

# 批量文件夹 -> 桌面时间戳新文件夹 OCR文字层_YYYYMMDD_HHMMSS
python scripts\pdf_ocr_overlay.py --input "D:\scans" --out-dir

# 指定页 / DPI
python scripts\pdf_ocr_overlay.py --input "x.pdf" --out "y.pdf" --pages 1,3-5 --dpi 300
```

输出 PDF 页面外观不变（文字层 render_mode=3 不可见），但文字可搜索、可复制。日志：`INPUTS` / `PAGE_OK` / `OK` / `FAIL` / `DONE`。

### 依赖安装（Python 3.12）

```powershell
python -m pip install pymupdf winrt-runtime winrt-Windows.Foundation winrt-Windows.Globalization winrt-Windows.Graphics.Imaging winrt-Windows.Media.Ocr winrt-Windows.Storage winrt-Windows.Storage.Streams
```

### 质量说明

- 标准印刷体（工程表格、宋体/黑体）识别准确；低质量/艺术字/竖排（如证件照、印章字）会出错，需人工核对。
- 输出体积会因嵌入中文字体略增（单页约 +3~5MB，多页共享字体）。

## 实现细节（供维护参考）

- **PDF→Word**：`batch_replace.py` 调用 `acrobat_export.ps1`，通过 `[System.__ComObject]::InvokeMember('saveAs', ...)` 走 Acrobat JS 桥接，静默导出 docx（与 GUI「导出到 Word」同一引擎）。中间 docx 落在临时目录，处理完自动清理。
- **文字替换**：用 `DispatchEx` 创建**独立** Word 实例，遍历 `doc.StoryRanges` 及 `NextStoryRange`（覆盖正文/页眉/页脚/文本框/脚注），对每个 story 执行 `Find.Execute` 的 wdReplaceAll。只关闭本次自己打开的文档，并只 Quit 这个自己创建的实例；用户已打开的 Word 绝不受影响。
- **踩坑**：`Find.Execute` 在 win32com 下必须用**完整位置参数**（11 个），只传命名参数 `Replace=2` 会静默不生效。脚本已按位置参数写好，勿改回命名参数。
- **OCR 踩坑**：WinRT 的 `OcrLine` 只有 `text`/`words`，坐标在 `OcrWord.bounding_rect`；winrt 4.x 的属性是 snake_case（`bounding_rect`、`pixel_width`）。文字层用 `render_mode=3` + 嵌入 `simhei.ttf`（比 msyh.ttc 体积小）。必须在同一 Python 环境同时具备 `fitz` 与 `winrt`（本机 PowerShell Python 3.12 已就绪）。
- **编码**：规则必须走 JSON 文件，不要命令行内联中文。
- **同名文件**：输出文件名冲突时自动追加 `_1`、`_2`。

## 资源

- `scripts/batch_replace.py`：主脚本（收集输入→PDF转换→Word替换→输出）。
- `scripts/acrobat_export.ps1`：Acrobat 静默转 Word 脚本（被主脚本调用）。
- `scripts/pdf_ocr_overlay.py`：扫描件 OCR 文字层脚本（Windows OCR + PyMuPDF）。
