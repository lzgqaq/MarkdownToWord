# Markdown 转 Word 使用说明

## 快速开始

1. 将压缩包完整解压到任意可写文件夹。
2. 双击 `MarkdownToWord/MarkdownToWord.exe`。
3. 点击“选择文件”，选中 `.md` 或 `.markdown` 文档。
4. 默认保存在源文件旁边，也可以点击“另存为…”选择位置。
5. 点击“转换为 Word”，完成后点击“打开 Word”或“打开文件夹”。

适用于 Windows 10/11 64 位。程序已包含 Python 运行组件和 Pandoc，无需另外安装，无需联网。查看生成的文件需使用 Word、WPS 或其他兼容 DOCX 的软件。

请保留 EXE 旁的 `_internal` 文件夹，不要单独移动 EXE。程序未进行商业代码签名；仅运行你确认来源的副本。

## 支持内容

- 一至六级标题、中文段落、粗体、斜体、删除线。
- 有序和无序列表，包含嵌套列表。
- 常见管道表格、可点击链接、引用。
- 行内代码、带底色的代码块。
- 本地图片，以 Markdown 所在目录为相对路径基准；按页面宽度缩放。
- UTF-8、UTF-8 BOM、带 BOM 的 UTF-16，以及 GB18030 中文文本。

默认输出 A4 页面，四边约 2 厘米边距，中文正文使用宋体，标题使用黑体，代码使用 Consolas。查看设备缺少相应字体时由办公软件替代。

文件已存在时会请求覆盖确认；转换失败不会提前清空旧文件。Word 正在占用目标文件时，请关闭该文档后重试。

## 图片和特殊语法

图片推荐使用 PNG 或 JPEG。图片路径含空格时，可使用 `![说明](<images/my picture.png>)`。

找不到的本地图片会显示文字占位并在转换完成后提示。网络图片不会自动下载；请先保存到本地并修改 Markdown 引用。普通网页超链接可以保留。

Mermaid 以代码显示，不会生成流程图。复杂 HTML、YAML 文档元数据和特殊 Markdown 扩展不作为第一版支持范围。数学公式由 Pandoc 尝试转换，本版本未对复杂公式进行专项验收。

这是内容和样式的转换工具，输出排版不会与所有 Markdown 编辑器的网页预览完全一致。

## 示例

`examples/示例.md` 与 `examples/示例.docx` 展示中文、六级标题、文本格式、列表、表格、引用、代码和本地图片。`examples/images` 是示例图片，请保留。

## 源码运行与重新打包

需要带 Tkinter 的 Python 3.10 或更新版本。在 `source` 目录打开终端：

```powershell
python -m pip install -r requirements.txt
python app.py
```

命令行转换：

```powershell
python converter.py "输入.md" "输出.docx"
python converter.py "输入.md" "输出.docx" --overwrite
```

重新构建 Windows 程序：

```powershell
python build.py
```

生成目录为 `dist/MarkdownToWord`。打包时需要 Windows 64 位 Python；不支持在其他系统直接构建 Windows EXE。

运行测试：

```powershell
python -m unittest -v test_converter
```

源代码中的 `patch_docx` 函数集中设置 Word 样式，可以修改字体、字号、页边距和表格颜色。

## 验证记录

本版本通过了五项自动测试：常见文档结构及图片尺寸、缺失与远程图片提示、覆盖与失败保护、GB18030、UTF-16 与无效输出路径。另对打包程序执行 Tk 界面初始化及实际转换检查。

示例 DOCX 已验证内部结构和图片嵌入。当前构建环境无法启动 Word 的页面预览，尚未完成逐页视觉检查；实际分页以你的 Word 或 WPS 为准。

## 第三方组件

Pandoc 的转换参数依据官方文档：https://pandoc.org/MANUAL.html 。第三方许可见 `licenses` 文件夹，版本及项目链接见 `第三方组件.md`。
