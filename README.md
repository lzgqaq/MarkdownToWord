# Markdown 转 Word

一个在 Windows 上本地运行的 Markdown 到 Word（DOCX）转换工具。

程序提供简单的桌面界面：选择 Markdown 文件，使用 Pandoc 生成可编辑的 Word 文档。文件不会上传到网络。

## 功能

- 一至六级标题、中文段落、粗体、斜体、删除线
- 有序/无序列表、嵌套列表、引用和表格
- 可点击链接、行内代码和代码块
- 插入 Markdown 文件旁边的本地图片
- UTF-8、UTF-8 BOM、UTF-16 和 GB18030 中文文本
- A4 页面、中文字体、标题层级、表格边框和图片宽度处理
- 覆盖确认、转换失败提示、打开 Word 和打开输出文件夹

## 直接使用

下载 [`release/MarkdownToWord-Windows.zip`](release/MarkdownToWord-Windows.zip)，完整解压后运行 `MarkdownToWord/MarkdownToWord.exe`。请保留 EXE 旁边的 `_internal` 文件夹。该 ZIP 包含程序、运行库、Pandoc、示例和使用说明。

输出文档可以使用 Microsoft Word、WPS 或其他支持 DOCX 的软件打开。

## 从源码运行

需要 Windows 64 位 Python 3.10 或更新版本，并带 Tkinter。

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r source/requirements.txt
python source/app.py
```

命令行转换：

```powershell
python source/converter.py "输入.md" "输出.docx"
python source/converter.py "输入.md" "输出.docx" --overwrite
```

运行测试：

```powershell
python -m unittest discover -s source -p "test_*.py" -v
```

## 构建 Windows 程序

```powershell
python source/build.py
```

程序会将 Pandoc 一起放入 `dist/MarkdownToWord`，因此使用者不需要另外安装 Pandoc。构建依赖见 `source/requirements.txt`。

## 图片路径

相对图片路径以 Markdown 文件所在目录为基准：

```markdown
![示例图片](images/sample.png)
```

缺失的本地图片会在文档中显示占位文字，并在转换完成后提示。网络图片不会自动下载。

## 限制

复杂 HTML、Mermaid 流程图和特殊 Markdown 扩展不保证还原。数学公式由 Pandoc 尝试转换，但未对复杂公式做专项处理。输出排版可能与 Markdown 编辑器的网页预览不同。

## 许可

本项目使用 MIT License。Pandoc、Python、PyInstaller 和其他第三方组件仍受各自许可证约束，相关文件在 `licenses` 目录。
