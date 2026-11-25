"""Local PDF OCR with native-PDF figure reconciliation, layout fidelity, and symbol verification."""

from io import BytesIO
from pathlib import Path
import re

from lxml import html as lxml_html
from PIL import Image


MATH_FACE = re.compile(r"(?:CMMI|CMSY|MSAM|MSBM|MathItalic|Math-Italic|MathSymbols|CambriaMath|Symbol)", re.I)
ROMAN_FACE = re.compile(r"(?:LMRoman|CMR\d|TimesNewRoman|Times-Roman)", re.I)
BOLD_FACE = re.compile(r"(?:Bold|Demi|Heavy|Black|CMBX|LMRoman\d*-Bold)", re.I)
ITALIC_FACE = re.compile(r"(?:Italic|Oblique|Slant|CMMI|Math-?Italic)", re.I)
HAN = r"\u3400-\u9fff"
SENTENCE_END = "。！？.!?:：；;”\"'）)]】》"
TOC_TITLES = {"目录", "contents", "tableofcontents"}
TOP_BACKMATTER = {"致谢", "参考文献", "references", "acknowledgements", "acknowledgments", "附录", "appendix"}
CAPTION_FIG = re.compile(r"^(?:图\s*\d|Fig\.?\s*\d|Figure\s*\d)", re.I)
CAPTION_TBL = re.compile(r"^(?:表\s*\d|Table\s*\d)", re.I)
CHAPTER_RE = re.compile(r"^(?:第\s*[0-9一二三四五六七八九十百零]+\s*章|Chapter\s+\w+)", re.I)
SECTION_RE = re.compile(r"^(\d+(?:\.\d+)+)(?!\d|\.)")
COVER_KV_RE = re.compile(r"^([^\s:：][^:：]{0,14})[:：]\s*(.+)$")
MATH_OPS = re.compile(r"[=<>±×÷≤≥≠≈∑∫√∞∂∇∈∉⊂⊃∪∩→←⇒⇔\u0370-\u03ff]")

# Content area inside A4 (210x297mm) with 25mm left/right and 23/22mm top/bottom margins.
MAX_IMG_WIDTH_MM = 158.0
MAX_IMG_HEIGHT_MM = 205.0
PT_TO_MM = 25.4 / 72.0


def inside(run, box):
    x0, y0, x1, y1 = run["bbox"]
    return (box[0] - 1 <= (x0 + x1) / 2 <= box[2] + 1
            and box[1] - 1 <= (y0 + y1) / 2 <= box[3] + 1)


def _overlap_ratio(inner, outer, pad=0.0):
    ix = max(0.0, min(inner[2], outer[2] + pad) - max(inner[0], outer[0] - pad))
    iy = max(0.0, min(inner[3], outer[3] + pad) - max(inner[1], outer[1] - pad))
    area = max(1.0, (inner[2] - inner[0]) * (inner[3] - inner[1]))
    return (ix * iy) / area


def _block_text(block):
    parts = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            parts.append(str(span.get("content", "")))
    return "".join(parts).strip()


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


def _native_runs_for_box(box, runs):
    """Return native runs whose center is inside box, or a single run enclosing box."""
    hits = [r for r in runs if inside(r, (box[0] - 2, box[1] - 2, box[2] + 2, box[3] + 2))]
    if hits:
        return hits, False
    # Check if box sits inside a single wider upright text run (e.g., a single CJK char cut out).
    for r in runs:
        rb = r["bbox"]
        if (rb[0] - 2 <= box[0] and box[2] <= rb[2] + 2
                and min(rb[3], box[3]) - max(rb[1], box[1]) > 0.5 * max(1.0, box[3] - box[1])):
            return [r], True
    return [], False


def _clean_inline_equation(span, runs, block, prev_span, next_span):
    """Demote false-positive OCR equations back to text and repair common OCR glitches."""
    content = span.get("content", "").strip()
    span["content"] = content
    if not content:
        return span

    # Citation superscript misread as math, e.g. 0.016^{[30]} -> 0.016<sup>[30]</sup>
    cite = re.fullmatch(r"([0-9.\s]+)\^\s*\{\s*\[\s*([\d\s]+)\s*\]\s*\}", content)
    if cite:
        num = re.sub(r"\s+", "", cite[1])
        ref = re.sub(r"\s+", "", cite[2])
        return {**span, "type": "text", "content": num, "_raw_sup": f"[{ref}]"}

    hits, enclosed = _native_runs_for_box(span["bbox"], runs)
    if not hits:
        return span

    if enclosed:
        # A single wider non-math run encloses this narrow span (e.g. '凸' inside '上基座受到凸轮曲线').
        run = hits[0]
        font = run["font"]["name"]
        if not MATH_FACE.search(font) and not ITALIC_FACE.search(font):
            rtxt = run["text"]
            prev_t = (prev_span.get("content", "") if prev_span and prev_span["type"] == "text" else "").strip()
            next_t = (next_span.get("content", "") if next_span and next_span["type"] == "text" else "").strip()
            # Find missing middle substring between prev tail and next head in the native run.
            recovered = ""
            if prev_t and next_t:
                p_anchor = prev_t[-min(4, len(prev_t)):]
                n_anchor = next_t[:min(4, len(next_t))]
                pi = rtxt.find(p_anchor)
                if pi >= 0:
                    ni = rtxt.find(n_anchor, pi + len(p_anchor))
                    if ni > pi + len(p_anchor) and ni - (pi + len(p_anchor)) <= 8:
                        recovered = rtxt[pi + len(p_anchor):ni]
            if recovered and not MATH_OPS.search(recovered):
                return {**span, "type": "text", "content": recovered}

    all_plain = all(
        not MATH_FACE.search(r["font"]["name"]) and not ITALIC_FACE.search(r["font"]["name"])
        for r in hits
    )
    sizes = [r["font"]["size"] for r in hits]
    uniform_size = (max(sizes) - min(sizes) <= 1.5) if sizes else False
    native_text = "".join(r["text"] for r in hits).strip()

    if all_plain and uniform_size and native_text and not MATH_OPS.search(native_text):
        next_t = (next_span.get("content", "") if next_span and next_span["type"] == "text" else "").lstrip()
        # Plain numbers/units/angles/ranges/CJK or single upright labels in captions/units.
        is_angle_or_unit = bool(re.fullmatch(
            r"[\d.\-+–—−]+(?:\s*°)?(?:\s*[-–—−]\s*[\d.]+\s*°)?(?:\s*[A-Za-z/%]+)?[,，。；;]?$",
            native_text
        ))
        is_cjk_plain = bool(re.search("[" + HAN + "]", native_text))
        is_caption_label = (block.get("original_label") == "figure_title"
                            and re.fullmatch(r"[A-Za-z0-9]{1,4}[,，。；;]?", native_text))
        is_trailing_unit = bool(re.fullmatch(r"[A-Za-z]{1,4}[,，。；;]?", native_text)
                                and (next_t.startswith(("。", "，", ",", "；", ";"))
                                     or re.search(r"_\s*\{\s*\\circ\s*\}|[,;；]$", content)))
        if is_angle_or_unit or is_cjk_plain or is_caption_label or is_trailing_unit:
            cleaned = native_text
            if cleaned and next_t and cleaned[-1] in ",，。；;" and next_t[0] in ",，。；;":
                cleaned = cleaned[:-1]
            return {**span, "type": "text", "content": cleaned}

    # Hallucinated parentheses around a single variable: '( f )' when native text is just 'f'.
    paren = re.fullmatch(r"\(\s*([A-Za-z]|\\[A-Za-z]+)\s*\)", content)
    if paren and "(" not in native_text and ")" not in native_text:
        span["content"] = paren[1]

    # Hallucinated \mathbf on regular upright (non-bold) native text, e.g. \mathbf{d}m -> \mathrm{d}m.
    if r"\mathbf" in span["content"] and not any(BOLD_FACE.search(r["font"]["name"]) for r in hits):
        span["content"] = span["content"].replace(r"\mathbf", r"\mathrm")

    return span


def _clean_interline_equation(span, runs):
    """Strip hallucinated trailing punctuation before \\tag{...} when absent in native PDF text."""
    content = span.get("content", "").strip()
    if re.search(r"\.\s*(?:\\tag\{|$)", content):
        hits, _ = _native_runs_for_box(span["bbox"], runs)
        native_text = "".join(r["text"].strip() for r in hits)
        if hits and not native_text.endswith("."):
            content = re.sub(r"\s*\.\s*(?=\\tag\{|$)", "", content)
    span["content"] = content


def _tag_line_pages(pages):
    """Track true source page index for lines merged across pages by RapidDoc's para_split."""
    n = len(pages)
    for page in pages:
        base_idx = page["page_idx"]
        for block in page.get("para_blocks", []):
            cur_idx = base_idx
            prev_y = None
            for line in block.get("lines", []):
                y0 = line["bbox"][1]
                if prev_y is not None and y0 < prev_y - 100 and cur_idx + 1 < n:
                    cur_idx += 1
                line["_page_idx"] = cur_idx
                prev_y = y0


def improve_text_layer(pages, native_pages):
    _tag_line_pages(pages)
    terms = set()
    for page in pages:
        for block in page["para_blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    if span["type"] in {"interline_equation", "inline_equation"}:
                        terms.update(word.replace(" ", "") for word in
                                     re.findall(r"\\mathrm\s*\{([A-Za-z ]+)\}", span["content"]))
    runs_by_page = []
    for native in native_pages:
        if native.get("rotation"):
            runs_by_page.append([])
        else:
            runs_by_page.append([
                r for b in native["blocks"] for l in b["lines"] for r in l["spans"]
                if r["text"].strip()
            ])

    for page in pages:
        for block in page["para_blocks"]:
            for line in block.get("lines", []):
                pidx = line.get("_page_idx", page["page_idx"])
                runs = runs_by_page[pidx] if 0 <= pidx < len(runs_by_page) else []
                spans = line["spans"]
                join_split_index(spans, runs)
                updated = []
                for i, span in enumerate(spans):
                    prev_s = spans[i - 1] if i > 0 else None
                    next_s = spans[i + 1] if i + 1 < len(spans) else None
                    if span["type"] in {"interline_equation", "inline_equation"}:
                        if not span.get("content", "").strip():
                            raise ValueError(f"PDF 第 {page['page_idx'] + 1} 页有无法识别的公式。")
                        repair_scripts(span, runs)
                        if span["type"] == "inline_equation":
                            span = _clean_inline_equation(span, runs, block, prev_s, next_s)
                        else:
                            _clean_interline_equation(span, runs)
                    if span["type"] == "inline_equation" and next_s is not None:
                        after = next_s.get("content", "").lstrip()
                        if after.startswith((",", "，")) and re.search(r"[=<>]", span["content"]):
                            span["content"] = re.sub(r"\s+\.$", "", span["content"])
                    if span["type"] == "text" and block["type"] == "text" and not span.get("_raw_sup"):
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

    repeated = {_block_text(b) for page in pages[1:] for b in page["discarded_blocks"]}
    for b in first["discarded_blocks"]:
        x0, y0, x1, y1 = b["bbox"]
        title = _block_text(b)
        if (b.get("original_label") == "header" and len(title) >= 4 and title not in repeated
                and height * .12 < y0 < height * .45 and y1 < top
                and abs((x0 + x1) / 2 - width / 2) < width * .08):
            blocks.insert(0, {**b, "type": "title", "original_label": "doc_title"})
            return


def _extract_native_figure_rects(pdf_page):
    """Collect non-full-page image XObjects and vector Form logos from a PDFium page."""
    import pypdfium2.raw as pdfium_c

    pw, ph = pdf_page.get_size()
    page_area = max(1.0, pw * ph)
    img_rects = []
    form_rects = []

    def bounds(obj):
        l, b, r, t = (pdfium_c.c_float(), pdfium_c.c_float(),
                      pdfium_c.c_float(), pdfium_c.c_float())
        if not pdfium_c.FPDFPageObj_GetBounds(obj, l, b, r, t):
            return None
        x0, y0, x1, y1 = l.value, ph - t.value, r.value, ph - b.value
        x0, y0 = max(0.0, x0), max(0.0, y0)
        x1, y1 = min(pw, x1), min(ph, y1)
        if x1 - x0 < 15 or y1 - y0 < 15:
            return None
        return (x0, y0, x1, y1)

    def has_paths(form_obj):
        cnt = pdfium_c.FPDFFormObj_CountObjects(form_obj)
        for j in range(cnt):
            sub = pdfium_c.FPDFFormObj_GetObject(form_obj, j)
            st = pdfium_c.FPDFPageObj_GetType(sub)
            if st in (pdfium_c.FPDF_PAGEOBJ_PATH, pdfium_c.FPDF_PAGEOBJ_IMAGE):
                return True
            if st == pdfium_c.FPDF_PAGEOBJ_FORM and has_paths(sub):
                return True
        return False

    def walk(obj, top_level=False):
        otype = pdfium_c.FPDFPageObj_GetType(obj)
        if otype == pdfium_c.FPDF_PAGEOBJ_IMAGE:
            rect = bounds(obj)
            if rect and (rect[2] - rect[0]) * (rect[3] - rect[1]) < 0.85 * page_area:
                img_rects.append(rect)
        elif otype == pdfium_c.FPDF_PAGEOBJ_FORM:
            rect = bounds(obj) if top_level else None
            before = len(img_rects)
            cnt = pdfium_c.FPDFFormObj_CountObjects(obj)
            for j in range(cnt):
                walk(pdfium_c.FPDFFormObj_GetObject(obj, j), top_level=False)
            # Pure vector Form (no raster images inside, contains paths, reasonable size) -> e.g. logo.
            if (rect and len(img_rects) == before and has_paths(obj)
                    and (rect[2] - rect[0]) * (rect[3] - rect[1]) < 0.50 * page_area):
                form_rects.append(rect)

    for i in range(pdfium_c.FPDFPage_CountObjects(pdf_page.raw)):
        walk(pdfium_c.FPDFPage_GetObject(pdf_page.raw, i), top_level=True)

    # Merge touching/overlapping raster image tiles into unified rects.
    merged = []
    for rect in sorted(img_rects, key=lambda r: (r[1], r[0])):
        for idx, prev in enumerate(merged):
            if (rect[0] <= prev[2] + 3 and rect[2] >= prev[0] - 3
                    and rect[1] <= prev[3] + 3 and rect[3] >= prev[1] - 3):
                merged[idx] = (min(prev[0], rect[0]), min(prev[1], rect[1]),
                               max(prev[2], rect[2]), max(prev[3], rect[3]))
                break
        else:
            merged.append(rect)
    return merged, form_rects


def _trim_white_border(img: Image.Image, crop_pt, pad_px=4):
    """Trim uniform near-white outer margin and adjust physical size proportionally."""
    rgb = img.convert("RGB")
    w, h = rgb.size
    # Build mask of non-white pixels (any channel < 248) via fast point lookup on grayscale.
    gray = rgb.convert("L")
    mask = gray.point(lambda p: 255 if p < 248 else 0)
    bbox = mask.getbbox()
    if not bbox:
        return rgb, crop_pt
    left = max(0, bbox[0] - pad_px)
    top = max(0, bbox[1] - pad_px)
    right = min(w, bbox[2] + pad_px)
    bottom = min(h, bbox[3] + pad_px)
    if right - left < 16 or bottom - top < 16:
        return rgb, crop_pt
    x0, y0, x1, y1 = crop_pt
    sx, sy = (x1 - x0) / w, (y1 - y0) / h
    new_pt = (x0 + left * sx, y0 + top * sy, x0 + right * sx, y0 + bottom * sy)
    return rgb.crop((left, top, right, bottom)), new_pt


def _render_pdf_crop(pdf_page, rect_pt, images_dir: Path, name: str):
    """Render a PDF region at crisp DPI, trim outer whitespace, and save with exact DPI metadata."""
    pw, ph = pdf_page.get_size()
    x0 = max(0.0, rect_pt[0] - 2.0)
    y0 = max(0.0, rect_pt[1] - 2.0)
    x1 = min(pw, rect_pt[2] + 2.0)
    y1 = min(ph, rect_pt[3] + 2.0)
    w_pt, h_pt = x1 - x0, y1 - y0
    # Use 288 DPI for compact logos/diagrams and 216 DPI for full-width figures.
    scale = 4.0 if max(w_pt, h_pt) < 220 else 3.0
    # pdfium crop uses (left, bottom, right, top) margins from page edges in PDF points.
    pil = pdf_page.render(
        scale=scale,
        crop=(x0, ph - y1, pw - x1, y0),
    ).to_pil()
    pil, (nx0, ny0, nx1, ny1) = _trim_white_border(pil, (x0, y0, x1, y1))
    w_mm = (nx1 - nx0) * PT_TO_MM
    h_mm = (ny1 - ny0) * PT_TO_MM
    dpi = round(scale * 72)
    pil.save(images_dir / name, format="PNG", dpi=(dpi, dpi), optimize=True)
    return name, (nx0, ny0, nx1, ny1), (w_mm, h_mm)


def reconcile_pdf_figures(source: Path, pages: list, raw_images: dict, images_dir: Path):
    """Merge over-segmented figure fragments onto native PDF images and stamp exact physical sizes."""
    import pypdfium2 as pdfium

    images_dir.mkdir(parents=True, exist_ok=True)
    pdf_doc = pdfium.PdfDocument(str(source))
    fig_counter = 0
    try:
        for page in pages:
            pidx = page["page_idx"]
            pdf_page = pdf_doc[pidx]
            blocks = page.get("para_blocks", [])
            if not any(b["type"] == "image" for b in blocks):
                continue

            img_rects, form_rects = _extract_native_figure_rects(pdf_page)
            candidate_rects = list(img_rects)
            # Only include a vector Form if it actually overlaps a layout-detected image block.
            for frect in form_rects:
                if any(_overlap_ratio(b["bbox"], frect, pad=6.0) >= 0.45
                       for b in blocks if b["type"] == "image"):
                    candidate_rects.append(frect)

            claimed_indices = set()
            replacements = {}  # earliest_idx -> unified image block
            absorbed_indices = set()

            for rect in candidate_rects:
                matched = []
                for bi, b in enumerate(blocks):
                    if bi in claimed_indices or b["type"] != "image":
                        continue
                    bb = b["bbox"]
                    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
                    center_in = (rect[0] - 6 <= cx <= rect[2] + 6 and rect[1] - 6 <= cy <= rect[3] + 6)
                    if center_in or _overlap_ratio(bb, rect, pad=6.0) >= 0.45:
                        matched.append(bi)
                if not matched:
                    continue

                fig_counter += 1
                filename = f"pdf_fig_{pidx + 1:03d}_{fig_counter:03d}.png"
                saved_name, final_bbox, (w_mm, h_mm) = _render_pdf_crop(pdf_page, rect, images_dir, filename)
                first_idx = min(matched)
                for bi in matched:
                    claimed_indices.add(bi)
                    if bi != first_idx:
                        absorbed_indices.add(bi)

                # Absorb internal figure text / colorbars / sub-panel labels strictly inside the figure rect.
                for bi, b in enumerate(blocks):
                    if bi in claimed_indices or bi in absorbed_indices or b["type"] in {"image", "title", "table"}:
                        continue
                    txt = _block_text(b)
                    if CAPTION_FIG.match(txt) or CAPTION_TBL.match(txt):
                        continue
                    if _overlap_ratio(b["bbox"], rect, pad=5.0) >= 0.55:
                        absorbed_indices.add(bi)

                replacements[first_idx] = {
                    "type": "image",
                    "bbox": list(final_bbox),
                    "_width_mm": w_mm,
                    "_height_mm": h_mm,
                    "blocks": [{
                        "type": "image_body",
                        "bbox": list(final_bbox),
                        "lines": [{
                            "bbox": list(final_bbox),
                            "spans": [{"type": "image", "bbox": list(final_bbox), "image_path": saved_name}],
                        }],
                    }],
                }

            new_blocks = []
            for bi, b in enumerate(blocks):
                if bi in absorbed_indices:
                    continue
                if bi in replacements:
                    new_blocks.append(replacements[bi])
                else:
                    new_blocks.append(b)
            page["para_blocks"] = new_blocks

        # For any remaining non-reconciled image blocks (e.g. scanned PDFs or raw vector regions),
        # save with accurate DPI metadata and record exact PDF physical dimensions.
        for page in pages:
            for b in page.get("para_blocks", []):
                if b["type"] != "image" or "_width_mm" in b:
                    continue
                bb = b["bbox"]
                w_mm = max(1.0, (bb[2] - bb[0]) * PT_TO_MM)
                h_mm = max(1.0, (bb[3] - bb[1]) * PT_TO_MM)
                b["_width_mm"], b["_height_mm"] = w_mm, h_mm
                for sb in b.get("blocks", []):
                    if sb["type"] != "image_body":
                        continue
                    for line in sb.get("lines", []):
                        for span in line.get("spans", []):
                            ip = span.get("image_path", "")
                            if not ip:
                                continue
                            fname = Path(ip).name
                            if (images_dir / fname).is_file():
                                continue
                            data = raw_images.get(ip) or raw_images.get(fname)
                            if not data:
                                continue
                            im = Image.open(BytesIO(data))
                            dpi_x = max(72, round(im.width / max(0.1, (bb[2] - bb[0]) / 72.0)))
                            dpi_y = max(72, round(im.height / max(0.1, (bb[3] - bb[1]) / 72.0)))
                            im.save(images_dir / fname, format="PNG", dpi=(dpi_x, dpi_y))
    finally:
        pdf_doc.close()


def _normalize_single_column_order(pages):
    """Fix vertical inversions on single-column pages and group side-by-side short rows (e.g. signatures)."""
    for page in pages:
        pw, _ = page["page_size"]
        blocks = [b for b in page.get("para_blocks", [])
                  if b["type"] in {"image", "table", "interline_equation"} or _block_text(b)]
        if len(blocks) < 2:
            page["para_blocks"] = blocks
            continue

        # Detect single-column pages: majority of text blocks span > 55% of page width, or no two-column prose.
        wide_text = [b for b in blocks if b["type"] == "text" and (b["bbox"][2] - b["bbox"][0]) > pw * 0.55]
        narrow_multiline = [b for b in blocks if b["type"] == "text"
                            and (b["bbox"][2] - b["bbox"][0]) < pw * 0.45 and len(b.get("lines", [])) >= 3]
        if narrow_multiline and not wide_text:
            page["para_blocks"] = blocks
            continue

        # Sort any strictly vertically inverted pairs on single-column pages.
        # Group side-by-side short single-line text blocks (like signature & date rows on declaration pages).
        ordered = sorted(blocks, key=lambda b: (round(b["bbox"][1] / 6.0), b["bbox"][0]))
        # Only replace RapidDoc order where an actual vertical inversion occurred.
        has_inversion = any(
            blocks[i]["bbox"][1] > blocks[i + 1]["bbox"][3] + 2.0
            for i in range(len(blocks) - 1)
        )
        if has_inversion:
            blocks = ordered

        # Pair side-by-side short single-line text blocks into a 2-column signature row marker.
        merged = []
        i = 0
        while i < len(blocks):
            cur = blocks[i]
            if (i + 1 < len(blocks) and cur["type"] == "text" and blocks[i + 1]["type"] == "text"
                    and len(cur.get("lines", [])) == 1 and len(blocks[i + 1].get("lines", [])) == 1
                    and cur.get("original_label") != "figure_title"
                    and blocks[i + 1].get("original_label") != "figure_title"):
                a, b = cur["bbox"], blocks[i + 1]["bbox"]
                min_h = max(1.0, min(a[3] - a[1], b[3] - b[1]))
                v_overlap = min(a[3], b[3]) - max(a[1], b[1])
                if (v_overlap >= 0.55 * min_h and (a[2] - a[0]) < pw * 0.45
                        and (b[2] - b[0]) < pw * 0.45 and a[2] < b[0] - 10):
                    left_blk, right_blk = (cur, blocks[i + 1]) if a[0] <= b[0] else (blocks[i + 1], cur)
                    merged.append({
                        "type": "_sig_row",
                        "bbox": [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])],
                        "cells": [left_blk, right_blk],
                    })
                    i += 2
                    continue
            merged.append(cur)
            i += 1
        page["para_blocks"] = merged


def _merge_cross_page_paragraphs(pages):
    """Join body paragraphs split across page / batch boundaries when RapidDoc could not merge them."""
    for i in range(len(pages) - 1):
        b1s = pages[i]["para_blocks"]
        b2s = pages[i + 1]["para_blocks"]
        if not b1s or not b2s:
            continue
        last_b, first_b = b1s[-1], b2s[0]
        if (last_b["type"] != "text" or first_b["type"] != "text"
                or last_b.get("original_label") != "text"
                or first_b.get("original_label") != "text"):
            continue
        t1 = _block_text(last_b)
        t2 = _block_text(first_b)
        if not t1 or not t2 or t1[-1] in SENTENCE_END:
            continue
        if re.match(r"^(?:\[\d+\]|\d+\.(?!\d)|\d+[、）)]|[(（]\d+[)）]|图\s*\d|表\s*\d|式中[:：])", t2):
            continue
        # Check that the continuation paragraph is NOT first-line indented.
        body_left = _page_body_left(pages[i + 1])
        first_line_x0 = first_b["lines"][0]["bbox"][0] if first_b.get("lines") else first_b["bbox"][0]
        if first_line_x0 - body_left >= 10.0:
            continue
        last_b["lines"].extend(first_b["lines"])
        b2s.pop(0)


def _page_body_left(page):
    pw, _ = page["page_size"]
    xs = []
    for b in page.get("para_blocks", []):
        if b["type"] == "text" and b.get("original_label") in {"text", "abstract"}:
            lines = b.get("lines", [])
            if len(lines) >= 2:
                xs.extend(l["bbox"][0] for l in lines[1:])
            elif b["bbox"][2] - b["bbox"][0] > pw * 0.55:
                xs.append(b["bbox"][0])
    if xs:
        xs.sort()
        return xs[len(xs) // 2]
    return pw * 0.115


def _find_toc_location(pages):
    """Return (page_idx, block_idx) of a printed Table of Contents heading, or None."""
    for p in pages:
        blocks = p["para_blocks"]
        for bi, b in enumerate(blocks):
            if b["type"] != "title":
                continue
            norm = re.sub(r"\s+", "", _block_text(b)).casefold()
            if norm not in TOC_TITLES:
                continue
            # Confirm TOC entries follow on this page or next page.
            tail = blocks[bi + 1:]
            if any(nb.get("original_label") == "content" or nb["type"] == "index"
                   or _block_text(nb).count(".") >= 6 for nb in tail):
                return (p["page_idx"], bi)
    return None


def _clean_html_table(markup: str) -> str:
    """Promote first row to thead and merge dangling unit-only first-column rows."""
    try:
        root = lxml_html.fragment_fromstring(markup, create_parent=False)
    except Exception:
        return markup
    if root.tag != "table":
        return markup

    # Normalize spaces immediately inside $...$ so Pandoc parses inline math in table cells.
    def fix_dollars(text):
        if not text or "$" not in text:
            return text
        return re.sub(r"\$\s*([^$]+?)\s*\$", lambda m: "$" + m[1].strip() + "$", text)

    for elem in root.iter():
        if elem.text:
            elem.text = fix_dollars(elem.text)
        if elem.tail:
            elem.tail = fix_dollars(elem.tail)

    rows = root.xpath(".//tr")
    # Merge dangling unit-only rows like <tr><td>(mm)</td><td></td>...<td></td></tr> into previous row.
    i = 1
    while i < len(rows):
        cells = rows[i].xpath("./td|./th")
        prev_cells = rows[i - 1].xpath("./td|./th")
        if len(cells) >= 3 and len(cells) == len(prev_cells):
            texts = ["".join(c.itertext()).strip() for c in cells]
            if (texts[0] and all(not t for t in texts[1:])
                    and re.fullmatch(r"[(（][^()（）]{1,14}[)）]", texts[0])):
                prev_cells[0].text = (prev_cells[0].text or "") + texts[0]
                rows[i].getparent().remove(rows[i])
                rows.pop(i)
                continue
        i += 1

    # Promote first row to <thead> if no <thead> exists and table has >= 2 rows.
    if not root.xpath("./thead") and len(rows) >= 2:
        first_tr = rows[0]
        first_tr.getparent().remove(first_tr)
        for cell in first_tr.xpath("./td"):
            cell.tag = "th"
        thead = lxml_html.Element("thead")
        thead.append(first_tr)
        root.insert(0, thead)

    return lxml_html.tostring(root, encoding="unicode")


def _fit_image_width_mm(w_mm: float, h_mm: float) -> float:
    scale = min(1.0, MAX_IMG_WIDTH_MM / max(1.0, w_mm), MAX_IMG_HEIGHT_MM / max(1.0, h_mm))
    return round(w_mm * scale, 1)


def _div(style: str, body: str) -> str:
    return f'::: {{custom-style="{style}"}}\n{body}\n:::'


def _table_fence(kind: str, html_str: str) -> str:
    return f"```{{={kind}}}\n{html_str}\n```"


def _escape_leading_ordered_list(text: str) -> str:
    """Prevent Pandoc from randomly turning one '1. xxx' paragraph into an <ol> while '2.xxx' is <p>."""
    return re.sub(r"^(\d+)\.(\s+)", r"\1\\.\2", text)


def _format_spans_block(block):
    """Render a block via RapidDoc's merge_para_with_text, preserving raw superscripts cleanly."""
    from rapid_doc.backend.pipeline.pipeline_middle_json_mkcontent import merge_para_with_text

    # Temporarily place unique placeholders for _raw_sup citations so escaping doesn't break HTML <sup>.
    sup_tokens = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            if span.get("_raw_sup"):
                token = f"T2WSUP{len(sup_tokens)}END"
                sup_tokens.append((token, span["_raw_sup"]))
                span["content"] = span["content"] + token
    rendered = merge_para_with_text(block)
    for token, sup in sup_tokens:
        rendered = rendered.replace(token, f"<sup>{sup}</sup>")
    return rendered


def _emit_front_matter_page(page, is_cover: bool, img_prefix: str):
    """Format a front-matter page (cover, declaration, abstract) with clean non-TOC styles."""
    pw, ph = page["page_size"]
    body_left = _page_body_left(page)
    blocks = page["para_blocks"]
    out = []
    i = 0
    seen_cover_title = False
    while i < len(blocks):
        b = blocks[i]
        btype = b["type"]

        # Cover metadata key-value run (e.g. 姓名/学号/学院/专业/指导教师).
        if is_cover and btype == "text":
            kv_rows = []
            j = i
            while j < len(blocks) and blocks[j]["type"] == "text" and len(blocks[j].get("lines", [])) == 1:
                m = COVER_KV_RE.match(_block_text(blocks[j]))
                if not m:
                    break
                kv_rows.append((m[1].strip(), m[2].strip()))
                j += 1
            if len(kv_rows) >= 3:
                trs = "".join(
                    f"<tr><td>{k}：</td><td>{v}</td></tr>"
                    for k, v in kv_rows
                )
                out.append(_table_fence("t2w-table-meta", f"<table><tbody>{trs}</tbody></table>"))
                i = j
                continue

        # Consecutive 2-column signature rows.
        if btype == "_sig_row":
            sig_rows = []
            j = i
            while j < len(blocks) and blocks[j]["type"] == "_sig_row":
                c0 = _format_spans_block(blocks[j]["cells"][0])
                c1 = _format_spans_block(blocks[j]["cells"][1])
                sig_rows.append((c0, c1))
                j += 1
            trs = "".join(f"<tr><td>{a}</td><td>{c}</td></tr>" for a, c in sig_rows)
            out.append(_table_fence("t2w-table-sig", f"<table><tbody>{trs}</tbody></table>"))
            i = j
            continue

        if btype == "image":
            w_mm = _fit_image_width_mm(b.get("_width_mm", 120.0), b.get("_height_mm", 80.0))
            x0, y0, x1, _ = b["bbox"]
            is_top_left_logo = is_cover and x0 < pw * 0.22 and x1 < pw * 0.65 and y0 < ph * 0.22
            style = "T2WFigureLeft" if is_top_left_logo else "T2WFigure"
            for sb in b.get("blocks", []):
                if sb["type"] == "image_body":
                    for line in sb.get("lines", []):
                        for span in line.get("spans", []):
                            if span.get("image_path"):
                                out.append(_div(style, f"![]({img_prefix}/{span['image_path']}){{width={w_mm}mm}}"))
        elif btype == "title":
            txt = _format_spans_block(b).replace("-\n", "").replace("\n", " ").strip()
            if not txt:
                i += 1
                continue
            orig = b.get("original_label", "")
            if is_cover:
                style = "T2WFrontTitle" if not seen_cover_title else "T2WFrontSubtitle"
                seen_cover_title = True
            elif orig == "doc_title":
                style = "T2WFrontSubtitle"
            else:
                style = "T2WFrontHeading"
            out.append(_div(style, txt))
        elif btype == "text":
            txt = _format_spans_block(b)
            if not txt:
                i += 1
                continue
            bx0, _, bx1, _ = b["bbox"]
            centered = abs((bx0 + bx1) / 2 - pw / 2) < pw * 0.10 and (bx1 - bx0) < pw * 0.50
            if is_cover and centered and len(b.get("lines", [])) == 1:
                out.append(_div("T2WFrontCenter", txt))
            else:
                # Bold abstract keyword labels cleanly.
                txt = re.sub(r"^(关键词\s*[:：]|Keywords\s*:|Key\s+words\s*:)\s*",
                             lambda m: f"**{m[1].strip()}** ", txt, flags=re.I)
                first_x0 = b["lines"][0]["bbox"][0] if b.get("lines") else bx0
                indented = (first_x0 - body_left >= 10.0
                            or (b.get("original_label") == "abstract" and len(b.get("lines", [])) >= 2
                                and first_x0 - min(l["bbox"][0] for l in b["lines"][1:]) >= 10.0))
                style = "T2WBodyIndent" if indented else "T2WBodyPlain"
                out.append(_div(style, _escape_leading_ordered_list(txt)))
        i += 1
    return out


def _build_markdown(pages: list, img_prefix: str) -> str:
    toc_loc = _find_toc_location(pages)
    # Use rich non-TOC front-matter formatting when multiple pages precede the TOC
    # (or when page 1 is a full thesis cover rather than a 1-2 line LaTeX title).
    rich_front = False
    toc_page_idx = -1
    if toc_loc is not None:
        toc_page_idx = toc_loc[0]
        if toc_page_idx > 1 or (toc_page_idx == 1 and len(pages[0]["para_blocks"]) >= 3):
            rich_front = True

    chunks = []
    in_references = False

    for page in pages:
        pidx = page["page_idx"]
        pw, _ = page["page_size"]
        blocks = page["para_blocks"]
        if not blocks:
            continue

        if rich_front and pidx < toc_page_idx:
            if pidx > 0 and chunks:
                chunks.append(r"\clearpage")
            chunks.extend(_emit_front_matter_page(page, is_cover=(pidx == 0), img_prefix=img_prefix))
            continue

        body_left = _page_body_left(page)
        i = 0
        while i < len(blocks):
            b = blocks[i]
            btype = b["type"]
            orig = b.get("original_label", "")

            if btype == "_sig_row":
                sig_rows = []
                j = i
                while j < len(blocks) and blocks[j]["type"] == "_sig_row":
                    c0 = _format_spans_block(blocks[j]["cells"][0])
                    c1 = _format_spans_block(blocks[j]["cells"][1])
                    sig_rows.append((c0, c1))
                    j += 1
                trs = "".join(f"<tr><td>{a}</td><td>{c}</td></tr>" for a, c in sig_rows)
                chunks.append(_table_fence("t2w-table-sig", f"<table><tbody>{trs}</tbody></table>"))
                i = j
                continue

            if btype == "title":
                txt = _format_spans_block(b).replace("-\n", "").replace("\n", " ").strip()
                if not txt:
                    i += 1
                    continue
                norm = re.sub(r"\s+", "", txt).casefold()
                if norm in {"参考文献", "references"}:
                    in_references = True
                elif CHAPTER_RE.match(txt) or SECTION_RE.match(txt):
                    in_references = False
                # Determine heading level cleanly.
                if CHAPTER_RE.match(txt) or norm in TOC_TITLES or norm in TOP_BACKMATTER:
                    level = 1
                elif m := SECTION_RE.match(txt):
                    level = min(6, m[1].count(".") + 1)
                else:
                    level = max(1, min(4, b.get("level") or 1))
                bx0, _, bx1, _ = b["bbox"]
                centered = (abs((bx0 + bx1) / 2 - pw / 2) < pw * 0.08
                            and (bx1 - bx0) < pw * 0.72 and level == 1)
                attr = " {.t2w-center}" if centered and norm not in TOC_TITLES else ""
                chunks.append(f"{'#' * level} {txt}{attr}")

            elif btype == "interline_equation":
                if b.get("lines") and b["lines"][0].get("spans"):
                    span = b["lines"][0]["spans"][0]
                    if span.get("content", "").strip():
                        chunks.append(_format_spans_block(b))
                    elif span.get("image_path"):
                        w_mm = _fit_image_width_mm(b.get("_width_mm", 120.0), b.get("_height_mm", 40.0))
                        chunks.append(_div("T2WFigure", f"![]({img_prefix}/{span['image_path']}){{width={w_mm}mm}}"))

            elif btype == "image":
                w_mm = _fit_image_width_mm(b.get("_width_mm", 120.0), b.get("_height_mm", 80.0))
                for sb in b.get("blocks", []):
                    if sb["type"] == "image_body":
                        for line in sb.get("lines", []):
                            for span in line.get("spans", []):
                                if span.get("image_path"):
                                    chunks.append(_div(
                                        "T2WFigure",
                                        f"![]({img_prefix}/{span['image_path']}){{width={w_mm}mm}}",
                                    ))
                    elif sb["type"] == "image_caption":
                        cap = _format_spans_block(sb)
                        if cap:
                            chunks.append(_div("T2WFigureCaption", cap))

            elif btype == "table":
                for sb in b.get("blocks", []):
                    if sb["type"] == "table_caption":
                        cap = _format_spans_block(sb)
                        if cap:
                            chunks.append(_div("T2WTableCaption", cap))
                for sb in b.get("blocks", []):
                    if sb["type"] == "table_body":
                        for line in sb.get("lines", []):
                            for span in line.get("spans", []):
                                if span.get("type") != "table":
                                    continue
                                if span.get("html"):
                                    if span.get("text_fallback"):
                                        from rapid_doc.model.table.fallback import markdown_prose
                                        chunks.append(markdown_prose(span["html"]))
                                    else:
                                        chunks.append(_table_fence("t2w-table", _clean_html_table(span["html"])))
                                elif span.get("image_path"):
                                    w_mm = _fit_image_width_mm(b.get("_width_mm", 140.0), b.get("_height_mm", 80.0))
                                    chunks.append(_div(
                                        "T2WFigure",
                                        f"![]({img_prefix}/{span['image_path']}){{width={w_mm}mm}}",
                                    ))
                for sb in b.get("blocks", []):
                    if sb["type"] == "table_footnote":
                        fn = _format_spans_block(sb)
                        if fn:
                            chunks.append(_div("T2WBodyPlain", fn))

            elif btype in {"text", "list", "index"}:
                txt = _format_spans_block(b)
                if not txt:
                    i += 1
                    continue
                if orig == "content" or btype == "index":
                    # Leave raw printed TOC entries bare so manuscript.pdf_headings detects and replaces them.
                    chunks.append(txt)
                elif (orig == "figure_title"
                      or ((CAPTION_FIG.match(txt) or CAPTION_TBL.match(txt))
                          and len(txt) <= 95 and "。" not in txt
                          and not re.match(r"^(?:图|表)\s*[\d\-a-zA-Z，,、和\s]+\s*(?:为|中|所示|是|给出|对比|展示)", txt))):
                    if CAPTION_TBL.match(txt):
                        chunks.append(_div("T2WTableCaption", txt))
                    else:
                        chunks.append(_div("T2WFigureCaption", txt))
                elif orig == "reference_content" or (in_references and re.match(r"^\[\d+\]", txt)):
                    chunks.append(_div("T2WReference", txt))
                else:
                    bx0 = b["bbox"][0]
                    first_x0 = b["lines"][0]["bbox"][0] if b.get("lines") else bx0
                    indented = first_x0 - body_left >= 10.0
                    if not indented and len(b.get("lines", [])) >= 2:
                        rest_x0 = min(l["bbox"][0] for l in b["lines"][1:])
                        indented = first_x0 - rest_x0 >= 10.0
                    style = "T2WBodyIndent" if indented else "T2WBodyPlain"
                    chunks.append(_div(style, _escape_leading_ordered_list(txt)))

            i += 1

    return "\n\n".join(chunks)


def pdf_markdown(source: Path, work: Path, lang: str, progress):
    from pdftext.extraction import dictionary_output
    from rapid_doc import RapidDoc

    progress("正在本机识别 PDF 版面、正文和公式。")
    # Retain the native PDF table route. For raster tables, use the OCR model
    # distributed with RapidDoc so a new table layout also works offline.
    engine = RapidDoc(lang=lang, pdf_pages_batch=8, table_config={"model_type": "slanet_1m"})
    result = engine(source, f_dump_content_list=False)
    pages = result.middle_json["pdf_info"]
    if not pages:
        raise ValueError("未读取到 PDF 页面。")
    restore_title(pages)
    images = work / "images"
    progress(f"已识别 {len(pages)} 页，正在整理图片、表格与版面")
    reconcile_pdf_figures(source, pages, result.images, images)
    _normalize_single_column_order(pages)
    _merge_cross_page_paragraphs(pages)
    progress(f"已识别 {len(pages)} 页，正在用 PDF 文字层核对行内符号")
    native = dictionary_output(str(source))
    improve_text_layer(pages, native)
    markdown = _build_markdown(pages, "images")
    if not markdown.strip():
        raise ValueError("PDF 未识别到可转换内容。")
    target = work / "recognized.md"
    target.write_text(markdown, encoding="utf-8")
    return target
