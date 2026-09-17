import os
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile
from converter import convert, ConversionError, W

NS = {'w':W}

class ConversionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix='中文路径 ')
        self.addCleanup(self.tmp.cleanup)
        self.folder = Path(self.tmp.name)

    def test_real_document_structure(self):
        sample = Path(__file__).resolve().parent.parent / 'examples' / '示例.md'
        output = self.folder / '中文 结果.docx'
        self.assertEqual(convert(sample, output), [])
        with zipfile.ZipFile(output) as z:
            self.assertIsNone(z.testzip())
            root = ET.fromstring(z.read('word/document.xml'))
            text = ''.join(root.itertext())
            self.assertIn('你好', text)
            self.assertTrue(root.findall('.//w:tbl',NS))
            self.assertTrue(root.findall('.//w:hyperlink',NS))
            self.assertTrue(root.findall('.//w:b',NS))
            self.assertTrue(root.findall('.//w:strike',NS))
            self.assertTrue(root.findall('.//w:numPr',NS))
            self.assertTrue(any(n.startswith('word/media/') for n in z.namelist()))
            styles = [n.attrib.get('{'+W+'}val') for n in root.findall('.//w:pStyle',NS)]
            for level in range(1,7): self.assertIn('Heading'+str(level), styles)
            drawing = root.find('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
            self.assertLessEqual(int(drawing.attrib['cx']), 6121000)

    def test_missing_and_remote_images(self):
        source = self.folder / '图片.md'
        source.write_text('![a](missing.png)\n\n![b](https://invalid.example/image.png)',encoding='utf-8')
        output = self.folder / '图片.docx'
        warnings = convert(source, output)
        self.assertEqual(len(warnings),2)
        with zipfile.ZipFile(output) as z: self.assertIn('图片缺失', z.read('word/document.xml').decode())

    def test_overwrite_and_bad_input_preserve_destination(self):
        source = self.folder / '文档.md'
        source.write_text('# 标题',encoding='utf-8')
        output = self.folder / '文档.docx'
        output.write_bytes(b'keep existing file')
        with self.assertRaises(ConversionError): convert(source,output)
        self.assertEqual(output.read_bytes(),b'keep existing file')
        with self.assertRaises(ConversionError): convert(self.folder/'missing.md',output,True)
        self.assertEqual(output.read_bytes(),b'keep existing file')
        convert(source,output,True)
        self.assertTrue(zipfile.is_zipfile(output))

    def test_legacy_encoding(self):
        source = self.folder / '旧编码.md'
        source.write_bytes('# 中文编码'.encode('gb18030'))
        self.assertIn('GB18030',' '.join(convert(source,self.folder/'编码.docx')))

    def test_utf16_and_invalid_output(self):
        source = self.folder / 'unicode.md'
        source.write_bytes('# 中文 UTF16'.encode('utf-16'))
        convert(source,self.folder/'utf16.docx')
        with self.assertRaises(ConversionError): convert(source,self.folder/'bad.txt')

if __name__ == '__main__': unittest.main(verbosity=2)
