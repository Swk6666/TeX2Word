"""Compare the native formulas in the two supplied GRU case outputs, in reading order."""

import ctypes
from io import BytesIO
import os
from pathlib import Path
import posixpath
import re
import unittest
from zipfile import ZipFile

from lxml import etree
import olefile

from tex2word.word import check, mathtype, verify


def body_equations(path, sdk):
    sdk.MTXFormSetTranslator.argtypes = [ctypes.c_short, ctypes.c_char_p]
    sdk.MTXFormEqn.argtypes = [ctypes.c_short, ctypes.c_short, ctypes.c_void_p, ctypes.c_long,
                               ctypes.c_short, ctypes.c_short, ctypes.c_void_p, ctypes.c_long,
                               ctypes.c_char_p, ctypes.c_void_p]
    check(sdk.MTXFormReset(), "重置翻译器")
    check(sdk.MTXFormSetTranslator(0, b"MathML2 (namespace attr).tdl"), "设置 MathML 翻译器")
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    with ZipFile(path) as package:
        rels = etree.fromstring(package.read("word/_rels/document.xml.rels"))
        destinations = {rel.get("Id"): rel.get("Target") for rel in rels}
        doc = etree.fromstring(package.read("word/document.xml"))
        for node in doc.xpath('//*[local-name()="OLEObject" and @ProgID="Equation.DSMT4"]'):
            part = posixpath.join("word", destinations[node.get(rid)])
            with olefile.OleFileIO(BytesIO(package.read(part))) as compound:
                native = compound.openstream("Equation Native").read()
            header_length = int.from_bytes(native[:2], "little")
            payload = native[header_length:]
            output = ctypes.create_string_buffer(1 << 20)
            dimensions = ctypes.create_string_buffer(32)
            # SDK types -3 = local memory, 4 = MTEF, 7 = translator text.
            check(sdk.MTXFormEqn(-3, 4, payload, len(payload), -3, 7, output,
                                 len(output), b"", dimensions), "从原生公式导出 MathML")
            yield etree.fromstring(output.value)


def symbols(node):
    children = [symbols(child) for child in node if isinstance(child.tag, str)]
    if etree.QName(node).localname == "mfenced":
        # MathML permits both literal delimiters and <mfenced> with implicit ones.
        separators = re.sub(r"\s+", "", node.get("separators", ","))
        pieces = []
        for index, child in enumerate(children):
            if index and separators:
                pieces.append(separators[min(index - 1, len(separators) - 1)])
            pieces.append(child)
        return node.get("open", "(") + "".join(pieces) + node.get("close", ")")
    return re.sub(r"\s+", "", (node.text or "") + "".join(children))


def script_arguments(root):
    templates = {"msub", "msup", "msubsup", "mfrac", "munder", "mover", "msqrt", "mroot", "munderover"}
    return [(etree.QName(node).localname, [symbols(child) for child in node])
            for node in root.iter() if etree.QName(node).localname in templates]


@unittest.skipUnless(os.environ.get("TEX2WORD_GRU_OUTPUT"), "需要两份 GRU 案例的输出目录")
class GruOutputTests(unittest.TestCase):
    def test_formulas_match_across_inputs(self):
        folder = Path(os.environ["TEX2WORD_GRU_OUTPUT"])
        paths = [folder / f"GRU控制方法说明-{kind}.docx" for kind in ("TeX", "PDF")]
        with mathtype() as sdk:
            for path in paths:
                self.assertEqual(verify(path), 48)
            tex, pdf = [list(body_equations(path, sdk)) for path in paths]
        self.assertEqual((len(tex), len(pdf)), (48, 48))
        for index, (expected, actual) in enumerate(zip(tex, pdf, strict=True), 1):
            with self.subTest(equation=index):
                self.assertEqual(symbols(actual), symbols(expected))
                self.assertEqual(script_arguments(actual), script_arguments(expected))
