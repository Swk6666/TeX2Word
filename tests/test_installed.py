"""Optional end-to-end check using only the installed application's private runtime."""

import os
from pathlib import Path
from shutil import copyfile
import subprocess
from tempfile import TemporaryDirectory
import unittest

from pypdf import PdfReader

from tex2word.word import editing, mathtype, read_mathml, verify, word_document


@unittest.skipUnless(os.environ.get("TEX2WORD_INSTALL"), "需要已安装的 TeX2Word")
class InstalledAppTests(unittest.TestCase):
    def test_bilingual_tex_and_scanned_pdf_offline(self):
        app = Path(os.environ["TEX2WORD_INSTALL"]).resolve()
        self.assertTrue((app / "TeX2Word.exe").is_file())
        runtime = app / "runtime"
        windows = Path(os.environ["SystemRoot"])
        env = dict(os.environ, PATH=os.pathsep.join(map(str, (windows / "System32", windows))),
                   HTTP_PROXY="http://127.0.0.1:9", HTTPS_PROXY="http://127.0.0.1:9",
                   ALL_PROXY="http://127.0.0.1:9", NO_PROXY="",
                   PYPANDOC_PANDOC=str(runtime / "Lib/site-packages/pypandoc/files/pandoc.exe"))
        fixtures = Path(__file__).parent
        with TemporaryDirectory(prefix="TeX2Word 安装测试 ") as work:
            for name, count, lang in (("bilingual.tex", 5, "ch"), ("english-scan.pdf", 4, "en")):
                with self.subTest(name=name):
                    source = Path(work) / ("自选文稿 " + name)
                    copyfile(fixtures / name, source)
                    if source.suffix == ".pdf":
                        self.assertFalse(any(page.extract_text().strip() for page in PdfReader(source).pages))
                    output = source.with_suffix(".docx")
                    process = subprocess.run([str(runtime / "python.exe"), "-I", "-B", "-X", "utf8",
                                              "-m", "tex2word", str(source), "-o", str(output), "--lang", lang],
                                             cwd=work, env=env, capture_output=True, timeout=300)
                    self.assertEqual(process.returncode, 0, process.stderr.decode("utf-8", "replace")[-3000:])
                    self.assertEqual(verify(output), count)
                    if source.suffix == ".tex":
                        with word_document(output, True) as doc, mathtype() as sdk:
                            for index, xpath in (
                                (2, '//*[local-name()="mi" and @mathvariant="script" and text()="A"]'),
                                (3, '//*[local-name()="msqrt"]'),
                                (5, '//*[local-name()="mtable" and count(*[local-name()="mtr"])=2]'),
                            ):
                                with editing(doc.InlineShapes(index), sdk) as ptr:
                                    self.assertTrue(read_mathml(sdk, ptr).xpath(xpath))
