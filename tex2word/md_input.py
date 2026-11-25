"""Read and normalize Markdown files with LaTeX math for Pandoc + MathType conversion."""

from pathlib import Path
import re


FENCE = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})")
MATH_ENV = re.compile(
    r"^(?P<indent>[ \t]*)\\begin\{(?P<env>equation|alignat|align|gather|multline|eqnarray)(?P<star>\*?)\}"
)
BRACKET_OPEN = re.compile(r"^[ \t]*\\\[[ \t]*$")
BRACKET_CLOSE = re.compile(r"^[ \t]*\\\][ \t]*$")
HTML_TABLE_OPEN = re.compile(r"^[ \t]*<table\b", re.I)
HTML_TABLE_CLOSE = re.compile(r"</table\s*>", re.I)
INLINE_CODE_OR_TEXT = re.compile(r"(`+)(?:(?!\1).)*?\1", re.S)
INLINE_PAREN_MATH = re.compile(r"\\\((.+?)\\\)", re.S)
INLINE_BRACKET_MATH = re.compile(r"\\\[(.+?)\\\]")
SINGLE_DOLLAR = re.compile(
    r"(?<!\\)(?<!\$)\$(?!\$)(?:"
    r"(?P<tight>[^ \t\n$](?:[^\n$]*?[^ \t\n$\\])?)"
    r"|[ \t]+(?P<spaced>[^\n$]+?)[ \t]+"
    r")(?<!\\)\$(?!\$)"
)
MATH_CUE = re.compile(r"\\[A-Za-z]+|[=^_{}+\-*/<>∑∫√±×÷≠≤≥≈∞α-ωΑ-Ω]")
BARE_CITATION = re.compile(r"^\s*\d+(?:\s*[-–,，]\s*\d+)*\s*$")


def read_markdown(path: Path) -> str:
    data = path.read_bytes()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("gb18030")
    return normalize_markdown(text)


def _normalize_inline_segment(segment: str) -> str:
    def replace_paren(match):
        body = match.group(1).strip()
        if not body:
            return match.group(0)
        return f"${body}$"

    def replace_bracket(match):
        body = match.group(1).strip()
        if not body or BARE_CITATION.match(body):
            return match.group(0)
        if not MATH_CUE.search(body):
            return match.group(0)
        return f"$${body}$$"

    def replace_dollar(match):
        spaced = match.group("spaced")
        if spaced is None:
            return match.group(0)
        body = spaced.strip()
        if not body or not (MATH_CUE.search(body) or re.fullmatch(r"[A-Za-z]", body)):
            return match.group(0)
        return f"${body}$"

    segment = INLINE_PAREN_MATH.sub(replace_paren, segment)
    segment = INLINE_BRACKET_MATH.sub(replace_bracket, segment)
    segment = SINGLE_DOLLAR.sub(replace_dollar, segment)
    return segment


def _normalize_line(line: str) -> str:
    parts = []
    pos = 0
    for match in INLINE_CODE_OR_TEXT.finditer(line):
        if match.start() > pos:
            parts.append(_normalize_inline_segment(line[pos:match.start()]))
        parts.append(match.group(0))
        pos = match.end()
    if pos < len(line):
        parts.append(_normalize_inline_segment(line[pos:]))
    return "".join(parts)


def normalize_markdown(text: str) -> str:
    """Normalize common Markdown math conventions outside fenced and inline code.

    Supports:
    - Standard Pandoc `$...$` and `$$...$$`
    - Typora / Obsidian / LLM-style `\\(...\\)` and `\\[...\\]`
    - Standalone AMS math environments (`\\begin{equation}`, `\\begin{align}`, etc.)
    - Inline `$ expr $` with inner spaces when containing math cues
    """
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out = []
    i = 0
    n = len(lines)
    fence_char = None
    fence_len = 0
    in_dollar_block = False

    while i < n:
        line = lines[i]

        # Track fenced code blocks verbatim.
        if fence_char is not None:
            out.append(line)
            m = re.match(rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_len},}}[ \t]*$", line)
            if m:
                fence_char = None
                fence_len = 0
            i += 1
            continue

        if not in_dollar_block:
            fm = FENCE.match(line)
            if fm:
                fence_char = fm.group(1)[0]
                fence_len = len(fm.group(1))
                out.append(line)
                i += 1
                continue

        # Track multi-line $$ ... $$ blocks verbatim.
        stripped = line.strip()
        if stripped == "$$":
            in_dollar_block = not in_dollar_block
            out.append(line)
            i += 1
            continue
        if in_dollar_block:
            out.append(line)
            i += 1
            continue

        # Raw HTML <table>...</table> blocks: wrap as atomic t2w-table so Pandoc
        # preserves the full table structure and inline TeX math inside cells.
        if HTML_TABLE_OPEN.match(line):
            j = i
            block = []
            found = False
            while j < n:
                block.append(lines[j])
                if HTML_TABLE_CLOSE.search(lines[j]):
                    found = True
                    break
                j += 1
            if found:
                out.append("```{=t2w-table}")
                out.extend(block)
                out.append("```")
                i = j + 1
                continue

        # Multi-line \[ ... \] display math block.
        if BRACKET_OPEN.match(line):
            j = i + 1
            block = []
            found = False
            while j < n:
                if BRACKET_CLOSE.match(lines[j]):
                    found = True
                    break
                block.append(lines[j])
                j += 1
            if found:
                out.append("$$")
                out.extend(block)
                out.append("$$")
                i = j + 1
                continue

        # Standalone AMS math environment (\begin{equation}...\end{equation}, align, etc.).
        em = MATH_ENV.match(line)
        if em:
            end_pat = re.compile(rf"\\end\{{{em.group('env')}{re.escape(em.group('star'))}\}}")
            j = i
            block = []
            found = False
            while j < n:
                block.append(lines[j])
                if end_pat.search(lines[j]):
                    found = True
                    break
                j += 1
            if found:
                out.append("$$")
                out.extend(block)
                out.append("$$")
                i = j + 1
                continue

        out.append(_normalize_line(line))
        i += 1

    return "\n".join(out)
