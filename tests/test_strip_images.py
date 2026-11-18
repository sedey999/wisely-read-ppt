#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""无图版生成脚本的回归测试。

重点覆盖：
1. 正文图片被删除；
2. 页眉/页脚图片也被删除（旧版只处理 body 的 bug）；
3. word/media、word/embeddings、docProps/thumbnail 被清零；
4. 所有 story part 的文字保持一致；
5. 输出文件体积明显小于带图版。
"""
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image
from docx import Document

SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS))

from strip_images import _collect_story_texts, strip_docx_images  # noqa: E402


def make_doc_with_images(path: Path):
    body_img = path.with_name("body.png")
    header_img = path.with_name("header.png")
    # 随机噪声图让 body 图片足够大，用于验证“无图版体积应显著下降”。
    Image.frombytes("RGB", (800, 600), os.urandom(800 * 600 * 3)).save(body_img)
    Image.new("RGB", (200, 100), (200, 30, 30)).save(header_img)

    doc = Document()
    doc.add_paragraph("正文前文字")
    doc.add_picture(str(body_img), width=doc.sections[0].page_width / 4)
    doc.add_paragraph("正文后文字：标点，“引号”。")

    section = doc.sections[0]
    section.header.paragraphs[0].text = "页眉文字"
    section.header.paragraphs[0].add_run().add_picture(str(header_img), width=1000000)
    section.footer.paragraphs[0].text = "页脚文字"
    section.footer.paragraphs[0].add_run().add_picture(str(header_img), width=600000)
    doc.save(str(path))


class StripImagesTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp(prefix="strip_test_"))

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_removes_body_header_footer_images_and_media(self):
        src = self.tmp / "with_images.docx"
        dst = self.tmp / "without_images.docx"
        make_doc_with_images(src)

        result, removed, verified, diff = strip_docx_images(src, dst)
        self.assertTrue(verified, diff)
        self.assertGreaterEqual(removed, 3)

        with zipfile.ZipFile(dst) as z:
            names = z.namelist()
            self.assertFalse([n for n in names if n.startswith("word/media/")])
            self.assertFalse([n for n in names if n.startswith("word/embeddings/")])
            self.assertFalse([n for n in names if n.startswith("docProps/thumbnail")])
            for name in ("word/document.xml", "word/header1.xml", "word/footer1.xml"):
                xml = z.read(name).decode("utf-8", "ignore")
                self.assertNotIn("<w:drawing", xml, name)
                self.assertNotIn("a:blip", xml, name)
                self.assertNotIn("v:imagedata", xml, name)

        with zipfile.ZipFile(src) as z:
            a = _collect_story_texts(z)
        with zipfile.ZipFile(dst) as z:
            b = _collect_story_texts(z)
        self.assertEqual(a, b)

        self.assertLess(os.path.getsize(dst), os.path.getsize(src) * 0.6)

    def test_text_only_doc_is_unchanged(self):
        src = self.tmp / "text_only.docx"
        dst = self.tmp / "text_only_out.docx"
        doc = Document()
        doc.add_paragraph("只有文字，没有图片。")
        doc.save(str(src))

        result, removed, verified, diff = strip_docx_images(src, dst)
        self.assertTrue(verified, diff)
        self.assertEqual(removed, 0)
        with zipfile.ZipFile(dst) as z:
            self.assertFalse([n for n in z.namelist() if n.startswith("word/media/")])

    @unittest.skipUnless(shutil.which("soffice"), "需要 LibreOffice 才能验证 PDF 可见图片")
    def test_rendered_pdf_has_no_images(self):
        src = self.tmp / "with_images.docx"
        dst = self.tmp / "without_images.docx"
        make_doc_with_images(src)
        result, removed, verified, diff = strip_docx_images(src, dst)
        self.assertTrue(verified, diff)

        out = self.tmp / "render"
        out.mkdir()
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out), str(dst)],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=90,
        )
        pdfs = list(out.glob("*.pdf"))
        self.assertTrue(pdfs)
        try:
            import pypdfium2 as pdfium
        except Exception:
            self.skipTest("未安装 pypdfium2，无法检查 PDF 图片")
        doc = pdfium.PdfDocument(str(pdfs[0]))
        try:
            for i in range(len(doc)):
                page = doc[i]
                try:
                    imgs = [o for o in page.get_objects() if isinstance(o, pdfium.PdfImage)]
                    self.assertEqual(len(imgs), 0)
                finally:
                    page.close()
        finally:
            doc.close()


if __name__ == "__main__":
    unittest.main()
