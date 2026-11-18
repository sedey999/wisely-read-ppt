#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""parse.py 的路径/参数/workspace 保护回归测试。"""
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL_DIR / "scripts"
PARSE = SCRIPTS / "parse.py"
sys.path.insert(0, str(SCRIPTS))

import parse  # noqa: E402


class WorkspaceDetectionTest(unittest.TestCase):
    def test_openclaw_workspace_env_is_respected(self):
        with tempfile.TemporaryDirectory(prefix="ws_test_") as td:
            ws = Path(td) / "workspace"
            ws.mkdir()
            with patch.dict(os.environ, {"OPENCLAW_WORKSPACE": str(ws)}):
                self.assertEqual(parse.find_workspace_dir(), ws.resolve())

    def test_cwd_subdir_of_workspace_is_detected(self):
        with tempfile.TemporaryDirectory(prefix="ws_test_") as td:
            ws = Path(td) / "workspace"
            ws.mkdir()
            sub = ws / "sub"
            sub.mkdir()
            with patch.dict(os.environ, {}, clear=False):
                os.environ.pop("OPENCLAW_WORKSPACE", None)
                old = Path.cwd()
                try:
                    os.chdir(sub)
                    # 祖先目录中有名为 workspace 的目录，应能识别。
                    self.assertEqual(parse.find_workspace_dir(), ws.resolve())
                finally:
                    os.chdir(old)


class CliGuardTest(unittest.TestCase):
    def _run(self, *args):
        return subprocess.run(
            [sys.executable, str(PARSE), *args],
            capture_output=True,
            text=True,
            timeout=30,
        )

    def test_rejects_out_of_range_dpi(self):
        r = self._run("not-exist.pptx", "--dpi", "1000", "--no-install")
        self.assertEqual(r.returncode, 1)
        self.assertIn("--dpi 必须在 50–600 之间", r.stdout)

    def test_rejects_non_docx_output(self):
        r = self._run("not-exist.pptx", "--out", "out.pdf", "--dpi", "150", "--no-install")
        self.assertEqual(r.returncode, 1)
        self.assertIn("--out 必须是 .docx", r.stdout)

    @unittest.skipUnless(shutil.which("soffice"), "需要 LibreOffice 才能解析 PPTX")
    def test_parse_outputs_real_image_free_docx(self):
        with tempfile.TemporaryDirectory(prefix="parse_test_") as td:
            from pptx import Presentation
            from pptx.util import Inches
            from PIL import Image

            img = Path(td) / "pic.png"
            Image.new("RGB", (800, 400), (20, 120, 200)).save(img)
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            slide.shapes.title.text = "标题"
            box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(1))
            box.text_frame.text = "正文，标点：“引号”。"
            slide.shapes.add_picture(str(img), Inches(1), Inches(3.5), width=Inches(4))
            pptx = Path(td) / "sample.pptx"
            prs.save(str(pptx))

            out = Path(td) / "sample_parse.docx"
            r = subprocess.run(
                [
                    sys.executable,
                    str(PARSE),
                    str(pptx),
                    "--out",
                    str(out),
                    "--dpi",
                    "100",
                    "--no-install",
                ],
                capture_output=True,
                text=True,
                timeout=180,
            )
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
            noimg = Path(td) / "sample_parse_无图.docx"
            self.assertTrue(out.exists())
            self.assertTrue(noimg.exists())

            import zipfile

            with zipfile.ZipFile(out) as z:
                self.assertTrue([n for n in z.namelist() if n.startswith("word/media/")])
            with zipfile.ZipFile(noimg) as z:
                self.assertFalse([n for n in z.namelist() if n.startswith("word/media/")])
                self.assertFalse([n for n in z.namelist() if n.startswith("docProps/thumbnail")])
            self.assertLess(noimg.stat().st_size, out.stat().st_size)


if __name__ == "__main__":
    unittest.main()
