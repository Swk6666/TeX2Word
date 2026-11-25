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
from tex2word.md_input import normalize_markdown, read_markdown
from tex2word.pdf_input import (
    _build_markdown,
    _clean_html_table,
    _merge_cross_page_paragraphs,
    improve_text_layer,
    inline_symbols,
    join_split_index,
    repair_scripts,
)
from tex2word.word import editing, insert_mathml, mathtype, read_mathml, style_document, verify, word_document


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

    def test_bilingual_ellipse_and_piecewise_function(self):
        path = FIXTURE.with_name("bilingual.tex")
        with TemporaryDirectory() as folder:
            manuscript = Manuscript(path, "latex", path.read_text(encoding="utf-8"))
            manuscript.write(Path(folder) / "bilingual.docx")
            self.assertEqual([eq.display for eq in manuscript.equations], [False, False, True, False, True])
            self.assertIn('mathvariant="script">A</mi>', manuscript.equations[1].mathml)
            self.assertIn("<msqrt", manuscript.equations[2].mathml)
            self.assertIn("<mtable", manuscript.equations[4].mathml)
            self.assertEqual(manuscript.labels["eq:area"].number, "1")

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
        from tex2word.convert import _publish
        with TemporaryDirectory() as folder:
            p = Path(folder)
            out = p / "existing.docx"
            out.write_bytes(b"existing")
            with self.assertRaises(FileExistsError):
                convert(FIXTURE, out)
            self.assertEqual(out.read_bytes(), b"existing")
            staged = p / "staged.docx"
            staged.write_bytes(b"new-bytes")
            with self.assertRaises(FileExistsError):
                _publish(staged, out, overwrite=False)
            self.assertEqual(out.read_bytes(), b"existing")
            _publish(staged, out, overwrite=True)
            self.assertEqual(out.read_bytes(), b"new-bytes")
            self.assertEqual([f.name for f in p.glob(".tex2word-*")], [])

    def test_html_table_math_and_headings(self):
        cleaned = _clean_html_table(
            "<table><tbody><tr><td>步数</td><td>1</td><td>2</td></tr>"
            "<tr><td>压缩量</td><td>2.5</td><td>5</td></tr>"
            "<tr><td>(mm)</td><td></td><td></td></tr></tbody></table>"
        )
        self.assertIn("<thead>", cleaned)
        self.assertIn("压缩量(mm)", cleaned)
        self.assertEqual(cleaned.count("<tr"), 2)

        md = (
            "# 目 录\n\n第1章 绪论........1\n1.3 研究........2\n\n"
            "# 第 1 章 绪论 {.t2w-center}\n\n"
            "# 1.3研究内容与研究目标\n\n"
            "```{=t2w-table}\n"
            "<table><thead><tr><th>材料</th><th>密度$(\\mathrm{kg}/\\mathrm{m}^3)$</th></tr></thead>"
            "<tbody><tr><td>钢</td><td>7850</td></tr></tbody></table>\n"
            "```\n"
        )
        with TemporaryDirectory() as folder:
            p = Path(folder)
            src = p / "doc.md"
            src.write_text(md, encoding="utf-8")
            ms = Manuscript(src, "markdown")
            staged = p / "out.docx"
            ms.write(staged)
            style_document(staged, ms)
            self.assertTrue(ms.book)
            self.assertTrue(ms.has_toc)
            self.assertEqual(ms.centered_headings, {1})
            self.assertEqual(len(ms.equations), 1)
            doc = Document(staged)
            self.assertEqual(len(doc.tables), 1)
            self.assertEqual(doc.tables[0].cell(1, 0).text, "钢")
            headings = [(para.style.name, para.text) for para in doc.paragraphs
                        if para.style and para.style.name.startswith("Heading ")]
            self.assertEqual(headings, [("Heading 1", "第 1 章 绪论"), ("Heading 2", "1.3研究内容与研究目标")])

    def test_pdf_false_positive_equations_and_prose_not_caption(self):
        def run(text, face, size, x0, x1, y0=100, y1=112):
            return {"text": text, "font": {"name": face, "size": size}, "bbox": [x0, y0, x1, y1]}

        pages = [
            {
                "page_idx": 0,
                "page_size": [595, 841],
                "para_blocks": [
                    {
                        "type": "text",
                        "original_label": "text",
                        "bbox": [68, 100, 520, 140],
                        "lines": [
                            {
                                "bbox": [94, 100, 520, 112],
                                "spans": [
                                    {"type": "text", "bbox": [94, 100, 160, 112], "content": "图 3-4a 为角度"},
                                    {"type": "inline_equation", "bbox": [162, 100, 182, 112], "content": r"4 5 ^ { \circ } )"},
                                    {"type": "text", "bbox": [184, 100, 220, 112], "content": "及频率"},
                                    {"type": "inline_equation", "bbox": [222, 100, 232, 112], "content": "( f )"},
                                    {"type": "text", "bbox": [234, 100, 320, 112], "content": "和阻尼比"},
                                    {"type": "inline_equation", "bbox": [322, 100, 360, 112], "content": r"0 . 0 1 6 ^ { [ 3 0 ] }"},
                                    {"type": "text", "bbox": [362, 100, 520, 112], "content": "，对应于一个周期的"},
                                ],
                            }
                        ],
                    }
                ],
            },
            {
                "page_idx": 1,
                "page_size": [595, 841],
                "para_blocks": [
                    {
                        "type": "text",
                        "original_label": "text",
                        "bbox": [68, 76, 520, 90],
                        "lines": [
                            {
                                "bbox": [70, 76, 520, 90],
                                "spans": [{"type": "text", "bbox": [70, 76, 520, 90], "content": "1.7e-2 s）。"}],
                            }
                        ],
                    }
                ],
            },
        ]
        native = [
            {
                "rotation": 0,
                "blocks": [{"lines": [{"spans": [
                    run("45°", "TimesNewRomanPSMT", 12.0, 162, 182),
                    run("f", "TimesNewRomanPS-ItalicMT", 12.0, 222, 232),
                ]}]}],
            },
            {"rotation": 0, "blocks": []},
        ]
        _merge_cross_page_paragraphs(pages)
        self.assertEqual(len(pages[1]["para_blocks"]), 0)
        improve_text_layer(pages, native)
        md = _build_markdown(pages, "images")
        self.assertIn('custom-style="T2WBodyIndent"', md)
        self.assertNotIn("T2WFigureCaption", md)
        self.assertIn("45°", md)
        self.assertNotIn(r"4 5 ^", md)
        self.assertIn("$f$", md)
        self.assertNotIn("( f )", md)
        self.assertIn("0.016<sup>[30]</sup>", md)
        self.assertIn("1.7e-2 s", md)

    def test_markdown_math_tables_and_cross_references(self):
        sample = FIXTURE.with_name("sample.md")
        normalized = normalize_markdown(sample.read_text(encoding="utf-8"))
        # Code fences and bare citations \[1\] must stay untouched.
        self.assertIn(r'expr = r"\(a + b\) and $x^2$"', normalized)
        self.assertIn(r"\[1\]", normalized)
        self.assertIn("$E = mc^2$", normalized)
        self.assertIn(r"$\alpha + \beta = \gamma$", normalized)

        with TemporaryDirectory() as folder:
            p = Path(folder)
            staged = p / "sample.docx"
            ms = Manuscript(sample, "markdown", read_markdown(sample))
            ms.write(staged)
            style_document(staged, ms)
            # 3 inline in intro + 2 display + 4 inline in pipe table + 1 inline in HTML table = 10
            self.assertEqual(len(ms.equations), 10)
            self.assertTrue(ms.book)
            self.assertEqual(ms.labels["eq:hooke"].number, "1.1")
            self.assertEqual(ms.equations[3].number, "1-1")
            with ZipFile(staged) as package:
                xml = package.read("word/document.xml")
            self.assertIn(b"REF MTEquation5", xml)
            doc = Document(staged)
            self.assertEqual(len(doc.tables), 2)
            self.assertEqual(len(ms.table_kinds), 2)

    def test_pdf_model_cache_and_verified_download(self):
        import hashlib
        from http.server import BaseHTTPRequestHandler, HTTPServer
        from threading import Thread
        from unittest.mock import patch
        from tex2word import models

        payload_a = b"A" * 1_100_000
        payload_b = b"B" * 1_100_000
        sha_a = hashlib.sha256(payload_a).hexdigest()
        sha_b = hashlib.sha256(payload_b).hexdigest()
        specs = (
            models.ModelSpec("layout", "layout.onnx", "版面模型", "bundled/layout.onnx", "layout.onnx", sha_a),
            models.ModelSpec("formula", "formula.onnx", "公式模型", "bundled/formula.onnx", "formula.onnx", sha_b),
        )
        served = {"/layout.onnx": payload_a, "/formula.onnx": payload_b}

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                data = served.get(self.path)
                if data is None:
                    self.send_response(404)
                    self.end_headers()
                    return
                self.send_response(200)
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def log_message(self, format, *args):
                pass

        server = HTTPServer(("127.0.0.1", 0), Handler)
        port = server.server_address[1]
        thread = Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with TemporaryDirectory() as folder:
                p = Path(folder).resolve()
                cache_dir = p / "cache"
                empty_site = p / "site"
                empty_site.mkdir()
                logs = []
                with patch.object(models, "PDF_MODELS", specs), patch.dict(os.environ, {"TEX2WORD_MODEL_DIR": str(cache_dir)}):
                    # 1. Downloads both models on first call and verifies SHA-256.
                    paths = models.ensure_pdf_models(logs.append, site_packages=empty_site, base_url=f"http://127.0.0.1:{port}/")
                    self.assertEqual(paths["layout"], cache_dir / "layout.onnx")
                    self.assertEqual(paths["formula"], cache_dir / "formula.onnx")
                    self.assertTrue(any("正在首次下载" in m for m in logs))

                    # 2. Second call hits cache with zero network access.
                    logs.clear()
                    paths2 = models.ensure_pdf_models(logs.append, site_packages=empty_site, base_url="http://127.0.0.1:1/")
                    self.assertEqual(paths2, paths)
                    self.assertEqual(logs, [])

                    # 3. Corrupted remote payload fails SHA-256 and leaves no .part file.
                    (cache_dir / "formula.onnx").unlink()
                    served["/formula.onnx"] = b"X" * 1_100_000
                    with self.assertRaises(RuntimeError) as ctx:
                        models.ensure_pdf_models(logs.append, site_packages=empty_site, base_url=f"http://127.0.0.1:{port}/")
                    self.assertIn("SHA-256", str(ctx.exception))
                    self.assertFalse((cache_dir / "formula.onnx.part").exists())
        finally:
            server.shutdown()
            server.server_close()

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
