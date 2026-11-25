# TeX2Word

在 Windows 上把中文、英文或中英混排的 TeX 和 PDF 转为 `.docx`。正文可在 Word 中直接修改，双击行内或独立公式会打开桌面版 MathType。

## 使用

1. 双击 `Start.cmd`，通过“浏览”选择 `.tex` 主文件或 `.pdf` 文件。
2. 选择 Word 输出位置。PDF 默认使用“中文或中英混排”，纯英文也可选择“英文”。
3. 点击“开始转换”，完成后点击“打开结果”。

程序使用本机的 Microsoft Word 和 MathType 7 桌面版。本机已用 Word 16.0 和 MathType 7.8 验证。换电脑时，请先安装这两个桌面程序，并准备 Python 3.12 64 位或 `uv`。`Start.cmd` 会在首次使用时调用 `Setup.cmd` 创建项目专用环境。PDF 识别模型在首次使用时联网下载，之后在本机推理。文稿由本地处理。

## 输入与输出

- 有 TeX 源码时，选择主 `.tex`。请保留相对引用的子文件、图片和已有的 `.aux`。程序读取常见宏、数学环境与章节，使用 `.aux` 中的公式编号生成书签及交叉引用。主文件支持 UTF-8、GB18030 和显式 TeX 编码声明。
- 只有 PDF 时，直接选择 PDF。程序在本机识别版面、正文和数学表达式，对含文字层的 PDF 还会用字形、字号和位置核对部分行内符号。扫描件也可直接读取。用于投稿前，请对照原稿检查细小上下标、相近符号和多栏阅读顺序。Word 会按可编辑内容重新分页。
- 默认写入源文件旁边的 `文件名-tex-MathType.docx` 或 `文件名-pdf-MathType.docx`。每个公式写入后由 MathType 回读。保存前，程序逐个检查 `Equation.DSMT4` 对象、OLE 关系和 `Equation Native` 数据流。验证通过才会发布输出文件。

## 命令行

在项目目录的 PowerShell 中运行以下示例。

```powershell
.\Start.cmd "D:\论文\main.tex" -o "D:\论文\论文-公式版.docx"
.\Start.cmd "D:\论文\paper.pdf" --lang en -o "D:\论文\paper-editable.docx"
.\Start.cmd --verify "D:\论文\论文-公式版.docx"
```

`--force` 允许替换已有输出。`--date "2026 年 9 月 5 日"` 可复现旧稿用 `\today` 编译时的标题页日期。自定义 MathType 安装可用环境变量 `MATHTYPE_DLL` 指向与 Python 位数一致的 `MT6.dll`。

## 本机验收

| 输入 | 独立公式 | 行内公式 | 检验项目 |
| --- | ---: | ---: | --- |
| GRU 说明的 TeX 主文件 | 19 | 29 | 按章编号、10 处公式引用、上下标、单点与双点导数、多行对齐、中文大括号注释 |
| 同一说明的 7 页 PDF | 19 | 29 | 独立读取 PDF，包含行内单字母、跨识别框的下标和中文公式注释 |
| 英文积分与矩阵示例的 TeX | 2 | 2 | 分数、定积分、矩阵及公式编辑后的保存和重开 |
| 上述英文示例的纯图像 PDF | 2 | 2 | 从扫描页恢复标题、英文正文与四个公式对象 |

两份 GRU 成品位于本机 `output` 目录。用 MathType 从两份成品逐个导出公式后，48 对公式的符号序列和上下标结构均一致。另在每份成品的副本中加入了一个带下标的变量，在 Word 中保存并重新打开，MathType 均能读回修改。TeX 示例的标题页使用原 PDF 的编译日期。

运行以下命令可重复执行回归测试。设置环境变量后，测试还会在 Word 中创建、修改、保存并重新打开一个含 MathType 公式的临时文档。

```powershell
$env:TEX2WORD_OFFICE_TEST = '1'
$env:TEX2WORD_GRU_OUTPUT = (Resolve-Path .\output).Path
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## 实现与依赖

Pandoc 解析文档结构并生成用于公式导入的 MathML。RapidDoc 负责本地 PDF 版面分析和公式识别。生成器通过 Word COM 与 MathType 桌面 SDK 写入原生 OLE 对象。对于带注释的大括号，生成器使用 MathType 的 MTEF 水平大括号模板，使括号可随内容伸展。[Pandoc 文档](https://pandoc.org/MANUAL.html) · [RapidDoc](https://github.com/RapidAI/RapidDoc) · [MathType OLE 的参考实现](https://github.com/DIPPER12138/dsh-mathtype) · [MTEF 格式说明](https://rtf2latex2e.sourceforge.net/MTEF5.html)
