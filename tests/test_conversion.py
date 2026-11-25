import os
from pathlib import Path
from shutil import copyfile
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from docx import Document
from lxml import etree

from tex2word.braces import prepare
from tex2word.convert import convert
from tex2word.manuscript import Manuscript
from tex2word.pdf_input import inline_symbols, join_split_index, repair_scripts
from tex2word.word import editing, insert_mathml, mathtype, read_mathml, verify, word_document


FIXTURE = Path(__file__).with_name("english.tex")


class ConversionTests(unittest.TestCase):
    def test_pdf_glyph_evidence(self):
        def glyph(text, face, size, x):
            return {"text": text, "font": {"name": face, "size": size}, "bbox": [x, 0, x+2, 8]}
        runs = [glyph("u", "CMMI12", 12, 0), glyph("k", "CMMI8", 8, 3),
                glyph("−", "CMSY8", 8, 6), glyph("1", "CMR8", 8, 9)]
        spans = [{"type": "inline_equation", "content": "u _ { k } .", "bbox": [0, 0, 6, 8]},
                 {"type": "text", "content": "−1 这个符号", "bbox": [6, 0, 50, 8]}]
        join_split_index(spans, runs)
        self.assertEqual((spans[0]["content"], spans[1]["content"]), ("u_{k-1}", "这个符号"))
        span = {"content": "回调 k 的状态", "bbox": [0, 0, 10, 10], "type": "text"}
        self.assertEqual(inline_symbols(span, [runs[1]], set())[1]["type"], "inline_equation")
        span["content"] = "x^{ call }"
        repair_scripts(span, [glyph("cal", "LMRoman8-Regular", 8, 0)])
        self.assertIn(r"\mathrm{cal}", span["content"])

    def test_brace_slots_leave_actual_fractions_intact(self):
        math = etree.fromstring("<math><munder><munder><mi>x</mi><mo>⏟</mo></munder>"
                                "<mtext>校准</mtext></munder><mfrac><mn>1</mn><mn>2</mn></mfrac></math>")
        braces = prepare(math)
        self.assertEqual(len(braces), 1)
        self.assertFalse(braces[0].top)
        self.assertEqual("".join(math[-1].itertext()), "12")

    def test_structure_numbers_and_reference(self):
        with TemporaryDirectory() as folder:
            target = Path(folder) / "intermediate.docx"
            manuscript = Manuscript(FIXTURE, "latex", FIXTURE.read_text(encoding="utf-8"))
            manuscript.write(target)
            self.assertEqual([eq.number for eq in manuscript.equations], ["", "1", "", "2"])
            self.assertIn("<mfrac>", manuscript.equations[1].mathml)
            self.assertIn("<mtable", manuscript.equations[3].mathml)
            with ZipFile(target) as package:
                xml = package.read("word/document.xml")
            self.assertIn(b"REF MTEquation2", xml)

    def test_relative_input_and_nested_math(self):
        with TemporaryDirectory() as folder:
            p = Path(folder)
            (p / "section.tex").write_text(
                r"\section{结果}中文 $x_1$.\footnote{The note contains $y$.}", encoding="utf-8")
            tex = r"\documentclass{article}\begin{document}\input{section}\end{document}"
            (p / "主稿.tex").write_text(tex, encoding="utf-8")
            m = Manuscript(p / "主稿.tex", "latex", tex)
            m.write(p / "result.docx")
            self.assertEqual([e.tex for e in m.equations], ["x_1", "y"])
            self.assertIn("中文", "\n".join(b.text for b in Document(p / "result.docx").paragraphs))

    def test_existing_output_is_preserved(self):
        with TemporaryDirectory() as folder:
            out = Path(folder) / "existing.docx"
            out.write_bytes(b"existing")
            with self.assertRaises(FileExistsError):
                convert(FIXTURE, out)
            self.assertEqual(out.read_bytes(), b"existing")

    @unittest.skipUnless(os.environ.get("TEX2WORD_OFFICE_TEST") == "1", "需要本机 Word 和 MathType")
    def test_edit_save_reopen(self):
        with TemporaryDirectory() as folder:
            out = Path(folder) / "英文结果.docx"
            convert(FIXTURE, out)
            self.assertEqual(verify(out), 4)
            edited = Path(folder) / "edited.docx"
            copyfile(out, edited)
            formula = ('<math xmlns="http://www.w3.org/1998/Math/MathML">'
                       '<mo>+</mo><msup><mi>z</mi><mn>9</mn></msup></math>')
            with word_document(edited) as doc, mathtype() as sdk:
                with editing(doc.InlineShapes(1), sdk) as ptr:
                    insert_mathml(sdk, ptr, formula)
                doc.Save()
            with word_document(edited, True) as doc, mathtype() as sdk:
                with editing(doc.InlineShapes(1), sdk) as ptr:
                    root = read_mathml(sdk, ptr)
                    self.assertTrue(root.xpath('//*[local-name()="msup"]'))
                    self.assertEqual("".join("".join(root.itertext()).split()), "a>0+z9")
            self.assertEqual(verify(edited), 4)


if __name__ == "__main__":
    unittest.main()
