"""Local Markdown to DOCX conversion. Python 3.10+, Pandoc 3.x."""
from __future__ import annotations

import argparse
import copy
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from urllib.parse import unquote, urlparse
import xml.etree.ElementTree as ET
import zipfile

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W)
def q(name): return '{' + W + '}' + name

class ConversionError(Exception):
    pass

def find_pandoc():
    root = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    candidates = [root / 'pandoc.exe', root / 'bin' / 'pandoc.exe',
                  Path(__file__).resolve().parent.parent / 'bin' / 'pandoc.exe']
    found = shutil.which('pandoc')
    if found: candidates.append(Path(found))
    for candidate in candidates:
        if candidate.is_file(): return str(candidate)
    try:
        import pypandoc
        return pypandoc.get_pandoc_path()
    except (ImportError, OSError):
        raise ConversionError('未找到 Pandoc。请保留完整程序目录，或安装 requirements.txt 中的依赖。')

def run(args, data=None, cwd=None):
    try:
        result = subprocess.run(args, input=data, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, cwd=cwd, timeout=180,
            creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    except subprocess.TimeoutExpired:
        raise ConversionError('转换超过 180 秒，请检查文档大小及图片。')
    except OSError as exc:
        raise ConversionError(f'无法启动转换引擎：{exc}') from exc
    if result.returncode:
        raise ConversionError(result.stderr.decode('utf-8', 'replace').strip() or '转换引擎执行失败。')
    return result

def element(parent, tag, **attrs):
    child = parent.find(q(tag))
    if child is None: child = ET.SubElement(parent, q(tag))
    for key, value in attrs.items(): child.set(q(key), str(value))
    if tag == 'rFonts':
        for key in ('asciiTheme','hAnsiTheme','eastAsiaTheme','cstheme'):
            child.attrib.pop(q(key),None)
    return child

def patch_docx(source, target):
    """Keep Pandoc's relationships and style IDs, modify only selected OOXML."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/styles.xml':
                root = ET.fromstring(data)
                defaults = element(element(element(root, 'docDefaults'), 'rPrDefault'), 'rPr')
                element(defaults, 'rFonts', ascii='Calibri', hAnsi='Calibri', eastAsia='宋体')
                element(defaults, 'sz', val=22)
                for style in root.findall(q('style')):
                    sid = style.get(q('styleId'), '')
                    if sid in ('Normal', 'BodyText', 'FirstParagraph', 'Compact', 'BlockText') or sid.startswith('Heading'):
                        rp = element(style, 'rPr')
                        element(rp, 'rFonts', ascii='Calibri', hAnsi='Calibri', eastAsia='宋体')
                        element(rp, 'color', val='000000').attrib.pop(q('themeColor'), None)
                        pp = element(style, 'pPr')
                        element(pp, 'spacing', after=120, line=300, lineRule='auto')
                    if sid.startswith('Heading') and sid[-1:].isdigit():
                        level = int(sid[-1])
                        element(element(style, 'rPr'), 'rFonts', eastAsia='黑体', ascii='Calibri', hAnsi='Calibri')
                        element(element(style, 'rPr'), 'sz', val={1:36,2:30,3:26}.get(level,24))
                        element(element(style, 'pPr'), 'keepNext')
                    if sid == 'BlockText': element(element(style, 'pPr'), 'ind', left=420)
                    if sid in ('SourceCode', 'VerbatimChar'):
                        rp = element(style, 'rPr')
                        element(rp, 'rFonts', ascii='Consolas', hAnsi='Consolas', eastAsia='等线')
                        element(rp, 'sz', val=19)
                        if sid == 'SourceCode':
                            pp = element(style, 'pPr')
                            element(pp, 'shd', fill='F2F4F7', val='clear')
                            element(pp, 'spacing', after=100, line=240, lineRule='auto')
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            elif item.filename == 'word/document.xml':
                root = ET.fromstring(data)
                for section in root.iter(q('sectPr')):
                    element(section, 'pgSz', w=11906, h=16838)
                    element(section, 'pgMar', top=1134, bottom=1134, left=1134, right=1134, header=567, footer=567, gutter=0)
                for table in root.iter(q('tbl')):
                    borders = element(element(table, 'tblPr'), 'tblBorders')
                    for side in ('top','left','bottom','right','insideH','insideV'):
                        element(borders, side, val='single', sz=4, color='D9D9D9')
                    for cell in table.iter(q('tc')):
                        props = element(cell, 'tcPr')
                        element(props, 'vAlign', val='center')
                        margins = element(props, 'tcMar')
                        for side in ('top','left','bottom','right'): element(margins, side, w=90, type='dxa')
                    first = table.find(q('tr'))
                    if first is not None:
                        element(element(first, 'trPr'), 'tblHeader')
                        for cell in first.findall(q('tc')): element(element(cell, 'tcPr'), 'shd', fill='EAF0F6', val='clear')
                data = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            zout.writestr(item, data)

def prepare_ast(ast, folder, warnings):
    def walk(node):
        if isinstance(node, list): return [walk(x) for x in node]
        if not isinstance(node, dict): return node
        if node.get('t') == 'Image':
            address = node['c'][2][0]
            parsed = urlparse(address)
            if parsed.scheme and not (len(parsed.scheme) == 1 and address[1:2] == ':'):
                warnings.append(f'未加载非本地图片：{address}')
                return {'t':'Str', 'c':f'[图片未加载：{address}]'}
            path = Path(unquote(address))
            if not path.is_absolute(): path = folder / path
            if not path.is_file():
                warnings.append(f'找不到图片：{address}')
                return {'t':'Str', 'c':f'[图片缺失：{address}]'}
            node = copy.deepcopy(node)
            node['c'][2][0] = str(path.resolve())
            # Clear dimensions so Pandoc fits images within page width.
            node['c'][0][2] = [v for v in node['c'][0][2] if v[0] not in ('width','height')]
        if node.get('t') in ('RawBlock','RawInline'):
            warnings.append('发现原始 HTML 或特殊标记，Word 可能无法保留其排版。')
        if node.get('t') == 'CodeBlock' and 'mermaid' in node['c'][0][1]:
            warnings.append('Mermaid 流程图以代码显示，未渲染为图片。')
        return {key:walk(value) for key,value in node.items()}
    return walk(ast)

def convert(source, destination, overwrite=False):
    source, destination = Path(source).resolve(), Path(destination).resolve()
    if not source.is_file() or source.suffix.lower() not in ('.md', '.markdown'):
        raise ConversionError('请选择存在的 .md 或 .markdown 文件。')
    if destination.suffix.lower() != '.docx': raise ConversionError('输出文件扩展名必须是 .docx。')
    if not destination.parent.is_dir(): raise ConversionError('输出文件夹不存在。')
    if destination.exists() and not overwrite: raise ConversionError('输出文件已存在，请更换文件名或确认覆盖。')
    warnings = []
    raw = source.read_bytes()
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        content = raw.decode('utf-16')
    else:
        try: content = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            try: content = raw.decode('gb18030'); warnings.append('已按 GB18030 编码读取源文件。')
            except UnicodeDecodeError: raise ConversionError('无法识别文本编码，请将 Markdown 保存为 UTF-8。')
    pandoc = find_pandoc()
    fmt = 'markdown-yaml_metadata_block-raw_attribute+pipe_tables+strikeout+fenced_code_blocks+hard_line_breaks'
    parsed = run([pandoc, '-f', fmt, '-t', 'json'], content.encode('utf-8'), source.parent)
    ast = prepare_ast(json.loads(parsed.stdout), source.parent, warnings)
    # Ignore document metadata so Markdown cannot change local conversion options.
    ast['meta'] = {}
    try:
        with tempfile.TemporaryDirectory(prefix='.md2word-', dir=destination.parent) as tmp:
            tmp = Path(tmp)
            original = tmp / 'default.docx'
            reference = tmp / 'reference.docx'
            original.write_bytes(run([pandoc, '--print-default-data-file=reference.docx']).stdout)
            patch_docx(original, reference)
            generated = tmp / 'generated.docx'
            result = run([pandoc, '-f', 'json', '-t', 'docx', '--standalone',
                '--reference-doc', str(reference), '--resource-path', str(source.parent),
                '-o', str(generated)], json.dumps(ast, ensure_ascii=False).encode('utf-8'), source.parent)
            if result.stderr.strip(): warnings.append(result.stderr.decode('utf-8','replace').strip())
            final = tmp / 'final.docx'
            patch_docx(generated, final)
            if overwrite: os.replace(final, destination)
            elif os.name == 'nt': os.rename(final, destination)  # Windows fails if destination appeared meanwhile.
            else:
                os.link(final, destination)
                final.unlink()
    except PermissionError as exc:
        raise ConversionError('无法保存文件。请关闭已打开的 Word 文件，并检查文件夹写入权限。') from exc
    return list(dict.fromkeys(warnings))

def main():
    parser = argparse.ArgumentParser(description='将本地 Markdown 转换为 Word')
    parser.add_argument('source')
    parser.add_argument('output')
    parser.add_argument('--overwrite', action='store_true')
    args = parser.parse_args()
    try:
        warnings = convert(args.source, args.output, args.overwrite)
        print('转换完成：' + str(Path(args.output).resolve()))
        for warning in warnings: print('提示：' + warning)
    except (ConversionError, OSError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0

if __name__ == '__main__': raise SystemExit(main())
