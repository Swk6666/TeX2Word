"""Local PDF OCR, with text-layer evidence for small mathematical symbols."""

from pathlib import Path
import re


MATH_FACE = re.compile(r"(?:CMMI|CMSY|MSAM|MSBM|MathItalic|Math-Italic|MathSymbols)", re.I)
ROMAN_FACE = re.compile(r"(?:LMRoman|CMR\d|TimesNewRoman|Times-Roman)", re.I)
HAN = r"\u3400-\u9fff"


def inside(run, box):
    x0, y0, x1, y1 = run["bbox"]
    return (box[0] - 1 <= (x0 + x1) / 2 <= box[2] + 1
            and box[1] - 1 <= (y0 + y1) / 2 <= box[3] + 1)


def repair_scripts(span, runs):
    roman = {r["text"].strip() for r in runs if inside(r, span["bbox"])
             and ROMAN_FACE.search(r["font"]["name"])}

    def script(match):
        def token(word):
            value = word[0]
            if value in roman:
                return r"\mathrm{" + value + "}"
            # Accept a one-letter OCR duplication only with local, unique font evidence.
            candidates = {value[:i] + value[i+1:] for i in range(1, len(value))
                          if value[i] == value[i-1]} & roman
            if len(candidates) == 1:
                return r"\mathrm{" + candidates.pop() + "}"
            return value
        return match[1] + re.sub(r"[A-Za-z]{2,}", token, match[2]) + "}"

    value = re.sub(r"([_^]\s*\{)([A-Za-z ,0-9]+)\}", script, span["content"])
    value = re.sub(r"([_^]\s*)\{\s*([" + HAN + r"]+)\s*\}",
                   lambda m: m[1] + r"{\text{" + m[2] + "}}", value)
    # Some recognizers put a redundant final row break before \end{aligned}.
    span["content"] = re.sub(r"\\\\\s*(\\end\{(?:aligned|gathered|array|matrix)\})", r"\1", value)


def join_split_index(spans, runs):
    for eq, text in zip(spans, spans[1:]):
        if eq["type"] != "inline_equation" or text["type"] != "text":
            continue
        suffix = re.match(r"\s*([−+\-]\d+)(?=\s|[" + HAN + "])", text["content"])
        sub = re.fullmatch(r"\s*([a-zA-Z])\s*_\s*\{\s*([a-zA-Z0-9]+)\s*\}\s*\.?\s*",
                           eq["content"])
        if not (suffix and sub):
            continue
        box = eq["bbox"]
        right = text["bbox"][0] + 8 * len(suffix[1])
        nearby = [r for r in runs if inside(r, (box[0], box[1]-3, right, box[3]+6))]
        glyphs = "".join(r["text"].strip() for r in nearby).replace("−", "-")
        wanted = sub[1] + sub[2] + suffix[1].replace("−", "-")
        if glyphs == wanted and len(nearby) > 1 and all(
                r["font"]["size"] < nearby[0]["font"]["size"] for r in nearby[1:]):
            eq["content"] = sub[1] + "_{" + wanted[1:] + "}"
            text["content"] = text["content"][suffix.end():].lstrip()


def inline_symbols(span, runs, terms):
    """Recover isolated math-face variables and upright terms also used in formulas."""
    text = span["content"]
    matches = {}
    for run in runs:
        if not inside(run, span["bbox"]):
            continue
        word, font = run["text"].strip(), run["font"]["name"]
        if MATH_FACE.search(font) and re.fullmatch(r"[A-Za-z\u0370-\u03ff]", word):
            tex = word
        elif (word in terms and word.isascii() and word.islower() and 2 <= len(word) <= 6
              and ROMAN_FACE.search(font) and re.search("[" + HAN + "]", text)):
            tex = r"\mathrm{" + word + "}"
        else:
            continue
        hits = list(re.finditer(r"(?<![A-Za-z0-9_])" + re.escape(word) + r"(?![A-Za-z0-9_])", text))
        if len(hits) == 1:
            matches[hits[0].span()] = tex
    out, start = [], 0
    for (a, b), tex in sorted(matches.items()):
        if a > start:
            out.append({**span, "content": text[start:a]})
        out.append({**span, "type": "inline_equation", "content": tex})
        start = b
    if start < len(text):
        out.append({**span, "content": text[start:]})
    return out


def improve_text_layer(pages, native_pages):
    terms = set()
    for page in pages:
        for block in page["para_blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    if span["type"] in {"interline_equation", "inline_equation"}:
                        terms.update(word.replace(" ", "") for word in
                                     re.findall(r"\\mathrm\s*\{([A-Za-z ]+)\}", span["content"]))
    for page, native in zip(pages, native_pages, strict=True):
        if native.get("rotation"):
            continue
        runs = [r for b in native["blocks"] for l in b["lines"] for r in l["spans"]
                if r["text"].strip()]
        for block in page["para_blocks"]:
            for line in block.get("lines", []):
                spans = line["spans"]
                join_split_index(spans, runs)
                updated = []
                for i, span in enumerate(spans):
                    if span["type"] in {"interline_equation", "inline_equation"}:
                        if not span.get("content", "").strip():
                            raise ValueError(f"PDF 第 {page['page_idx'] + 1} 页有无法识别的公式。")
                        repair_scripts(span, runs)
                    if span["type"] == "inline_equation" and i + 1 < len(spans):
                        after = spans[i + 1].get("content", "").lstrip()
                        if after.startswith((",", "，")) and re.search(r"[=<>]", span["content"]):
                            span["content"] = re.sub(r"\s+\.$", "", span["content"])
                    if span["type"] == "text" and block["type"] == "text":
                        updated.extend(inline_symbols(span, runs, terms))
                    else:
                        updated.append(span)
                line["spans"] = updated


def restore_title(pages):
    """Keep a centered first-page title that the layout model classified as a header."""
    first = pages[0]
    blocks = first["para_blocks"]
    if not blocks or any(b.get("original_label") == "doc_title" for b in blocks):
        return
    width, height = first["page_size"]
    top = min(b["bbox"][1] for b in blocks)

    def content(block):
        return "".join(s.get("content", "") for line in block.get("lines", []) for s in line["spans"]).strip()

    repeated = {content(b) for page in pages[1:] for b in page["discarded_blocks"]}
    for b in first["discarded_blocks"]:
        x0, y0, x1, y1 = b["bbox"]
        title = content(b)
        if (b.get("original_label") == "header" and len(title) >= 4 and title not in repeated
                and height * .12 < y0 < height * .45 and y1 < top
                and abs((x0 + x1) / 2 - width / 2) < width * .08):
            blocks.insert(0, {**b, "type": "title", "original_label": "doc_title"})
            return


def pdf_markdown(source: Path, work: Path, lang: str, progress):
    from pdftext.extraction import dictionary_output
    from rapid_doc import RapidDoc
    from rapid_doc.backend.pipeline.pipeline_middle_json_mkcontent import union_make
    from rapid_doc.utils.enum_class import MakeMode

    progress("正在本机识别 PDF 版面、正文和公式。")
    # Retain the native PDF table route. For raster tables, use the OCR model
    # distributed with RapidDoc so a new table layout also works offline.
    engine = RapidDoc(lang=lang, pdf_pages_batch=8, table_config={"model_type": "slanet_1m"})
    result = engine(source, f_dump_content_list=False)
    pages = result.middle_json["pdf_info"]
    if not pages:
        raise ValueError("未读取到 PDF 页面。")
    restore_title(pages)
    progress(f"已识别 {len(pages)} 页，正在用 PDF 文字层核对行内符号")
    native = dictionary_output(str(source))
    improve_text_layer(pages, native)
    markdown = union_make(pages, MakeMode.MM_MD, "images")
    if not markdown.strip():
        raise ValueError("PDF 未识别到可转换内容。")
    images = work / "images"
    images.mkdir(exist_ok=True)
    for name, data in result.images.items():
        (images / Path(name).name).write_bytes(data)
    target = work / "recognized.md"
    target.write_text(markdown, encoding="utf-8")
    return target
