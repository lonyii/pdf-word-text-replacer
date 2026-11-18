# pdf-word-text-replacer

一个 [WorkBuddy](https://www.workbuddy.cn/) 技能：批量把 **PDF / Word** 文件转成可编辑 Word、按规则查找替换文字、**给扫描件 PDF 叠加 OCR 隐藏文字层**，再把结果另存到桌面上一个以日期时间戳命名的新文件夹。

PDF 的转换走 **Adobe Acrobat Pro** 自带的导出引擎（与菜单「文件 → 导出到 → Microsoft Word」完全相同的引擎），但通过 IAC/JS 接口静默调用——**无需鼠标、不弹对话框、不抢焦点**。文字替换走 Microsoft Word COM，覆盖正文、页眉、页脚、表格、文本框，并保留原格式。扫描件 OCR 走 **Windows 内置中文 OCR**，完全静默可批量。

## 功能

- 📂 输入一个**文件夹**、**一批 PDF**、**一批 Word**，或三者混合（`.pdf` / `.docx` / `.doc`）；文件夹会递归扫描子目录。
- 🔄 PDF 自动转成可编辑 Word（Acrobat 静默导出，无 GUI 操作）。
- ✏️ 按多组规则查找替换文字，保留原始格式。
- 🔍 扫描件 PDF 叠加**隐藏 OCR 文字层**：页面外观不变，文字变得可搜索、可复制（Windows 中文 OCR，静默）。
- 📁 结果统一另存到桌面的 `Word替换_YYYYMMDD_HHMMSS` / `OCR文字层_YYYYMMDD_HHMMSS` 新文件夹（也可自定义输出目录）。
- 🛡️ 只处理本次打开的文件，不影响用户已经打开的 Word / Acrobat 窗口。

## 目录结构

```
pdf-word-text-replacer/
├── SKILL.md                      # WorkBuddy 技能定义
├── README.md
├── scripts/
│   ├── batch_replace.py          # 主脚本：收集输入 → PDF转换 → Word替换 → 输出
│   ├── acrobat_export.ps1        # Acrobat 静默导出 PDF→Word（被主脚本调用）
│   └── pdf_ocr_overlay.py        # 扫描件 OCR 文字层（Windows OCR + PyMuPDF）
└── examples/
    └── rules.example.json        # 替换规则示例
```

## 安装为 WorkBuddy 技能

把整个仓库克隆（或下载）到用户级技能目录：

```bash
git clone https://github.com/lonyii/pdf-word-text-replacer.git \
  "%USERPROFILE%\.workbuddy\skills\pdf-word-text-replacer"
```

重启 WorkBuddy 后即可使用。

## 环境依赖

- Windows 系统
- PDF→Word / 文字替换：**Adobe Acrobat Pro**（非 Reader）+ **Microsoft Word** + `pywin32`
- 扫描件 OCR：Windows 内置中文 OCR（系统自带）+ `pymupdf` + WinRT 模块

```bash
python -m pip install pywin32
# 扫描件 OCR 额外依赖：
python -m pip install pymupdf winrt-runtime winrt-Windows.Foundation winrt-Windows.Globalization winrt-Windows.Graphics.Imaging winrt-Windows.Media.Ocr winrt-Windows.Storage winrt-Windows.Storage.Streams
```

## 在 WorkBuddy 中使用

直接对 WorkBuddy 描述需求即可，例如：

> 把 `E:\工程资料` 文件夹里的 PDF 和 Word 都转成 Word，把里面的「甲方」替换成「建设单位」、「NC26E444」替换成「NC99XXXX」，输出到桌面。

扫描件 OCR：

> 把 `D:\扫描件` 文件夹里的扫描 PDF 做 OCR，加上隐藏文字层，输出到桌面新文件夹。

WorkBuddy 会自动写规则文件、调用脚本并核对结果。

## 扫描件 OCR（加隐藏文字层）

对纯扫描的图片型 PDF，用 `pdf_ocr_overlay.py` 叠加不可见文字层（页面外观不变，但文字可搜索/可复制）：

```powershell
# 单个文件 -> 同名 _ocr.pdf
python scripts\pdf_ocr_overlay.py --input "D:\scans\扫描件.pdf"

# 整个文件夹 -> 桌面 OCR文字层_时间戳 文件夹
python scripts\pdf_ocr_overlay.py --input "D:\scans" --out-dir

# 自定义输出 / 只处理指定页
python scripts\pdf_ocr_overlay.py --input "x.pdf" --out "y.pdf" --pages 1,3-5 --dpi 300
```

**为什么不用 Acrobat 自带 OCR**：Acrobat 的"识别文本"是受限命令，无法通过 IAC 静默脚本真正执行（命令可触发但不产生文字层；官方方式需写 Program Files 受信任脚本 + 管理员 + 重启，且语言对话框无法静默设置）。Windows 内置 OCR 完全静默、可批量、中文标准印刷体识别准确。

## 命令行用法

### 1. 准备规则文件

**务必把规则写成 UTF-8 的 `.json 文件`**，不要在命令行内联中文 JSON（PowerShell 5.1 传给 Python 时会破坏中文编码）。格式见 [`examples/rules.example.json`](examples/rules.example.json)：

```json
[
  { "find": "甲方", "replace": "建设单位" },
  { "find": "乙方", "replace": "施工单位", "whole_word": false }
]
```

每条规则字段：

| 字段 | 必填 | 说明 |
|---|---|---|
| `find` | 是 | 要查找的文本 |
| `replace` | 否 | 替换成的文本（默认空串，即删除） |
| `match_case` | 否 | 区分大小写，默认 `false` |
| `whole_word` | 否 | 全字匹配，默认 `false` |
| `wildcards` | 否 | 使用通配符，默认 `false` |

### 2. 运行脚本

```powershell
python scripts\batch_replace.py `
  --input "E:\工程资料" "D:\其他\文件.pdf" `
  --rules "examples\rules.example.json"
```

参数：

| 参数 | 说明 |
|---|---|
| `--input` | 一个或多个路径：文件 / 文件夹 / 通配符，可混合 |
| `--rules` | 规则 JSON 文件路径（推荐），也接受 JSON 字符串（仅纯 ASCII 时） |
| `--out-dir` | 可选，自定义输出目录；默认是桌面时间戳文件夹 |
| `--no-convert-pdf` | 可选，跳过 PDF 转换（只处理 Word 文件） |

### 3. 输出

脚本打印以 `|` 分隔的日志：

```
RULES|[...]
INPUTS|3|["...", "..."]
OUTDIR|C:\Users\...\Desktop\Word替换_20261118_150800
CONVERT_OK|<pdf>|<临时docx>
REPLACE_OK|<源>|<输出docx>|<字节数>
DONE|3|C:\Users\...\Desktop\Word替换_20261118_150800
```

## 实现说明

- **PDF→Word**：`acrobat_export.ps1` 通过 `[System.__ComObject]::InvokeMember('saveAs', ...)` 调用 Acrobat 的 JS 桥接对象 `saveAs(path, "com.adobe.acrobat.docx")`，与 GUI 导出同一引擎。脚本会关闭自己打开的文档，但绝不退出 Acrobat 进程。
- **文字替换**：`batch_replace.py` 用 `DispatchEx` 创建独立的 Word 实例，遍历 `doc.StoryRanges` 及 `NextStoryRange` 覆盖所有故事（正文/页眉/页脚/文本框/脚注），对每个 story 执行 `Find.Execute` 的 `wdReplaceAll`。结束后只关闭本次打开的文档、只退出自己创建的实例。
- **注意**：`Find.Execute` 在 win32com 下必须用**完整的 11 个位置参数**，只传命名参数 `Replace=2` 会静默不生效。
- **OCR 文字层**：`pdf_ocr_overlay.py` 逐页以 300 DPI 渲染 → Windows `Windows.Media.Ocr`（zh-Hans-CN）识别每个词及其像素坐标 → 按比例映射回 PDF 坐标 → 用 PyMuPDF `insert_text(render_mode=3)` 以嵌入的 `simhei.ttf` 写入不可见文字。标准印刷体识别准确；低质量/艺术字需人工核对。

## 许可证

私有仓库，仅供个人使用。
