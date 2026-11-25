"""Keep Pandoc's document structure and place Word fields beside equation anchors."""

import json
import re
from pathlib import Path
from uuid import uuid4
from xml.sax.saxutils import escape

from .mathml import Equation, make_mathml, pandoc
from .tex_input import expand_inputs


ENV = re.compile(r"^\s*\\begin\{(equation|alignat|align|gather|multline|eqnarray)(\*?)\}(.*)\\end\{\1\2\}\s*$", re.S)
LABEL = re.compile(r"\\label\s*\{([^{}]+)\}")
TAG = re.compile(r"\\tag(\*?)\s*\{([^{}]+)\}")
REF = re.compile(r"\\(eqref|ref)\s*\{([^{}]+)\}")
BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def raw(kind, xml):
    return {"t": "Raw" + kind, "c": ["openxml", xml]}


def text_run(text):
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def field(instruction, result):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {escape(instruction)} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + text_run(result) + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def split_alignment(body):
    """Only split top-level AMS rows. A matrix or cases inside a row stays intact."""
    pieces, start, depth = [], 0, 0
    for token in re.finditer(r"\\begin\{[^}]+\}|\\end\{[^}]+\}|\\\\(?:\[[^]]*\])?", body):
        if token[0].startswith(r"\begin"):
            depth += 1
        elif token[0].startswith(r"\end"):
            depth -= 1
        elif depth == 0:
            pieces.append(body[start:token.start()])
            start = token.end()
    return [s for s in [*pieces, body[start:]] if s.strip()]


class Manuscript:
    def __init__(self, source: Path, kind: str, tex: str = "", date: str | None = None):
        self.source = source
        self.kind = kind
        self.tex = tex = expand_inputs(source, tex) if kind == "latex" else tex
        self.book = bool(re.search(r"\\documentclass(?:\[[^\]]*\])?\{(?:ctex)?(?:book|report)\}", tex)
                         or re.search(r"\\chapter\*?\s*\{", tex))
        self.zh = bool(re.search(r"[\u3400-\u9fff]", tex))
        self.prefix = "MTX" + uuid4().hex[:12] + "Z"
        self.chapter = self.seq = 0
        self.equations = []
        self.labels = {}
        self.warnings = set()
        self.has_toc = False
        self.date = date
        self.aux = {}
        aux_path = source.with_suffix(".aux")
        if kind == "latex" and aux_path.is_file():
            self.aux = dict(re.findall(r"\\newlabel\{([^{}]+)\}\{\{([^{}]+)\}",
                                       aux_path.read_text("utf-8", errors="replace")))

    def equations_for(self, math):
        mode, latex = math["c"]
        display = mode["t"] == "DisplayMath"
        env = ENV.match(latex) if display else None
        name, star, body = env.groups() if env else ("", "", latex)
        if name == "alignat":
            body = re.sub(r"^\s*\{\d+\}", "", body)
        rows = split_alignment(body) if name in {"align", "alignat", "gather", "eqnarray"} else [body]
        for row in rows:
            labels = LABEL.findall(row)
            tag = TAG.search(row)
            no_number = re.search(r"\\(?:notag|nonumber)\b", row)
            number = ""
            if name and not star and not no_number:
                self.seq += 1
                number = f"{self.chapter}.{self.seq}" if self.book and self.chapter else str(self.seq)
            if labels and labels[0] in self.aux:
                number = self.aux[labels[0]]
            if tag:
                number = tag[2]
            expr = re.sub(r"\\(?:notag|nonumber)\b", "", TAG.sub("", LABEL.sub("", row))).strip()
            if name in {"align", "alignat", "eqnarray"}:
                expr = r"\begin{aligned}" + expr + r"\end{aligned}"
            elif name == "multline":
                expr = r"\begin{gathered}" + expr + r"\end{gathered}"
            if not expr:
                raise ValueError(f"遇到空公式：{latex[:120]}")
            index = len(self.equations) + 1
            eq = Equation(f"{self.prefix}{index:06d}", expr, display, number,
                          f"MTEquation{index}" if labels and number else "")
            if eq.bookmark:
                for label in labels:
                    if label in self.labels:
                        raise ValueError(f"公式标签重复：{label}")
                    self.labels[label] = eq
            self.equations.append(eq)
            yield eq, bool(tag and tag[1])

    def display_block(self, eq, literal_tag=False):
        # A4 and 25 mm margins, matching the section set in word.py.
        fmt = '<w:jc w:val="center"/>'
        content = text_run(eq.marker)
        if eq.number:
            fmt = ('<w:tabs><w:tab w:val="center" w:pos="4535"/>'
                   '<w:tab w:val="right" w:pos="9070"/></w:tabs>')
            serial = 100000 + len(self.equations)
            number = text_run(eq.number)
            if eq.bookmark:
                number = (f'<w:bookmarkStart w:id="{serial}" w:name="{eq.bookmark}"/>'
                          + number + f'<w:bookmarkEnd w:id="{serial}"/>')
            if not literal_tag:
                number = text_run("(") + number + text_run(")")
            tab = "<w:r><w:tab/></w:r>"
            content = tab + content + tab + number
        return raw("Block", '<w:p><w:pPr>' + fmt +
                   '<w:spacing w:before="65" w:after="70" w:line="240" w:lineRule="auto"/>'
                   '</w:pPr>' + content + '</w:p>')

    def paragraph(self, block):
        blocks, current = [], []

        def flush():
            while current and current[-1]["t"] in {"Space", "SoftBreak"}:
                current.pop()
            if current:
                blocks.append({"t": block["t"], "c": current.copy()})
            current.clear()

        for node in block["c"]:
            if node["t"] == "Math" and node["c"][0]["t"] == "DisplayMath":
                flush()
                for eq, literal in self.equations_for(node):
                    blocks.append(self.display_block(eq, literal))
            elif node["t"] not in {"Space", "SoftBreak"} or current:
                current.append(self.rewrite(node))
        flush()
        return blocks

    def rewrite(self, value):
        if isinstance(value, list):
            result = []
            for item in value:
                if isinstance(item, dict) and item.get("t") in {"Plain", "Para"}:
                    result.extend(self.paragraph(item))
                else:
                    result.append(self.rewrite(item))
            return result
        if not isinstance(value, dict):
            return value
        kind, data = value.get("t"), value.get("c")
        if kind == "Math":
            eq, _ = next(self.equations_for(value))
            return raw("Inline", text_run(eq.marker))
        if self.book and kind == "Header" and data[0] == 1 and "unnumbered" not in data[1][1]:
            self.chapter += 1
            self.seq = 0
        if kind == "RawBlock" and data[0] in {"tex", "latex"}:
            cmd = data[1].strip()
            if cmd == r"\tableofcontents":
                self.has_toc = True
                title = "目录" if self.zh else "Contents"
                return raw("Block", '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/>'
                           '<w:pageBreakBefore/><w:keepNext/></w:pPr>' + text_run(title) +
                           '</w:p><w:p>' + field(r'TOC \o "1-3" \h \z', "") + '</w:p>')
            if cmd in {r"\clearpage", r"\newpage", r"\cleardoublepage"}:
                return raw("Block", BREAK)
            if cmd != r"\maketitle":
                self.warnings.add(f"未排版的 TeX 指令：{cmd[:100]}")
            return raw("Block", "")
        return {key: self.rewrite(item) for key, item in value.items()}

    def references(self, value):
        if isinstance(value, list):
            return [self.references(node) for node in value]
        if not isinstance(value, dict):
            return value
        data = value.get("c")
        if value.get("t") == "RawInline" and data[0] in {"tex", "latex"}:
            ref = REF.fullmatch(data[1])
            if ref:
                eq = self.labels.get(ref[2])
                if eq:
                    code = field(f"REF {eq.bookmark} \\h", eq.number)
                    if ref[1] == "eqref":
                        code = text_run("(") + code + text_run(")")
                    return raw("Inline", code)
                self.warnings.add(f"未解析的引用：{ref[2]}")
                return {"t": "Str", "c": self.aux.get(ref[2], f"[{ref[2]}]")}
            self.warnings.add(f"未排版的 TeX 指令：{data[1][:100]}")
        return {key: self.references(item) for key, item in value.items()}

    def write(self, target):
        args = ["-f", self.kind, "-t", "json", self.source.name]
        if self.kind == "latex":
            args[1] = "latex+raw_tex"
            args[-1] = "-"
            tree = json.loads(pandoc(args, self.source.parent, self.tex.encode("utf-8")))
        else:
            tree = json.loads(pandoc(args, self.source.parent))
            self.pdf_headings(tree)
        if self.date is not None:
            tree["meta"]["date"] = {"t": "MetaString", "c": self.date}
        tree = self.references(self.rewrite(tree))
        make_mathml(self.equations, tree["pandoc-api-version"], self.source.parent)
        args = ["-f", "json", "-t", "docx", "-o", str(target)]
        if self.kind == "latex":
            args.append("--number-sections")
        pandoc(args, self.source.parent, json.dumps(tree).encode())

    def pdf_headings(self, tree):
        """Use a live Word TOC when OCR recovered a printed table of contents."""
        blocks = tree["blocks"]

        def heading(item):
            if item["t"] != "Header":
                return ""
            return "".join(n.get("c", "") if n["t"] == "Str" else " " for n in item["c"][2]).strip()

        for i, block in enumerate(blocks):
            title = heading(block)
            if title.casefold() not in {"目录", "目 录", "contents", "table of contents"}:
                continue
            following = next((j for j in range(i + 1, len(blocks)) if heading(blocks[j])), None)
            if following is None:
                continue
            entries = blocks[i+1:following]
            if not entries or any(b["t"] not in {"Plain", "Para"} for b in entries):
                continue
            dots = sum(n.get("c", "").count(".") for b in entries for n in b["c"] if n["t"] == "Str")
            if dots < 6:
                continue
            self.zh = bool(re.search(r"[\u3400-\u9fff]", title))
            blocks[i:following] = [{"t": "RawBlock", "c": ["latex", r"\tableofcontents"]}]
            if i in (1, 2) and heading(blocks[0]):
                date = []
                if i == 2 and blocks[1]["t"] in {"Para", "Plain"}:
                    date = blocks[1]["c"]
                date_text = " ".join(n["c"] for n in date if n["t"] == "Str")
                if i == 1 or (len(date_text) < 45 and re.search(r"\b\d{4}\b", date_text)):
                    tree["meta"]["title"] = {"t": "MetaInlines", "c": blocks[0]["c"][2]}
                    if date:
                        tree["meta"]["date"] = {"t": "MetaInlines", "c": date}
                    del blocks[:i]
            break
        for block in blocks:
            title = heading(block)
            if re.match(r"第\S+章|Chapter\s+\w+", title, re.I):
                self.book = True
            if match := re.match(r"^(\d+(?:\.\d+)+)\b", title):
                block["c"][0] = min(6, match[1].count(".") + 1)
