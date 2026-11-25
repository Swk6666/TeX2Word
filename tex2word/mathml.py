from dataclasses import dataclass, field
import json
from pathlib import Path
import subprocess

import pypandoc
from lxml import etree, html

from .braces import Brace, prepare


@dataclass
class Equation:
    marker: str
    tex: str
    display: bool
    number: str = ""
    bookmark: str = ""
    mathml: str = ""
    braces: list[Brace] = field(default_factory=list)


def pandoc(args: list[str], cwd: Path, data: bytes | None = None) -> bytes:
    result = subprocess.run(
        [pypandoc.get_pandoc_path(), *args], input=data, cwd=cwd,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    messages = result.stderr.decode("utf-8", "replace").strip()
    if result.returncode:
        raise RuntimeError(messages or f"Pandoc 退出码 {result.returncode}")
    if "Could not fetch resource" in messages:
        raise ValueError(messages)
    if messages:
        print(messages, flush=True)
    return result.stdout


def make_mathml(equations: list[Equation], version: list[int], cwd: Path) -> None:
    """Use Pandoc's TeX parser, including structures the MathType TeX importer drops."""
    if not equations:
        return
    blocks = [
        {"t": "Para", "c": [{"t": "Math", "c": [
            {"t": "DisplayMath" if eq.display else "InlineMath"}, eq.tex
        ]}]} for eq in equations
    ]
    doc = {"pandoc-api-version": version, "meta": {}, "blocks": blocks}
    result = pandoc(["-f", "json", "-t", "html5", "--mathml", "--wrap=none"],
                    cwd, json.dumps(doc).encode())
    fragment = html.fragment_fromstring(result.decode("utf-8"), create_parent=True)
    paragraphs = fragment.findall("p")
    if len(paragraphs) != len(equations):
        raise ValueError("公式解析结果与输入数量不一致。")
    for eq, paragraph in zip(equations, paragraphs):
        presentation = paragraph.find("math/semantics")
        if presentation is None or not len(presentation) or presentation[0].tag == "annotation":
            raise ValueError(f"公式无法转换为 MathML：{eq.tex[:180]}")
        math = etree.Element("math", {"xmlns": "http://www.w3.org/1998/Math/MathML",
                                   "display": "block" if eq.display else "inline"})
        node = presentation[0]
        node.tail = None
        for table in node.iter("mtable"):
            row = table.find("mtr")
            if row is not None and any(c.get("columnalign") for c in row):
                align = {c.get("columnalign", "center") for c in row}
                # The desktop editor supports one horizontal alignment per matrix.
                table.set("columnalign", align.pop() if len(align) == 1 else "left")
        math.append(node)
        eq.braces = prepare(math)
        # The desktop SDK accepts UTF-16 but some translators still use a system
        # code page internally. Numeric XML references preserve CJK and math symbols.
        eq.mathml = etree.tostring(math, encoding="unicode").encode("ascii", "xmlcharrefreplace").decode()
