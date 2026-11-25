"""Bridge stretchy TeX braces to MathType's native horizontal-brace template.

The desktop MathML importer maps U+23DF to a missing character in a limits
template. Let MathType itself encode the two slots, including fonts, scripts
and CJK, then give these slots to the MTEF 5 horizontal-brace template.
"""

from dataclasses import dataclass
import re
from uuid import uuid4

from lxml import etree


@dataclass(frozen=True)
class Brace:
    key: str
    top: bool
    level: int


def prepare(math):
    braces = []
    for inner in reversed(list(math.iter("munder", "mover"))):
        if len(inner) != 2 or inner[1].tag != "mo" or inner[1].text not in {"⏟", "⏞", "︸", "︷"}:
            continue
        top = inner.tag == "mover"
        outer = inner.getparent()
        decorated = outer.tag == inner.tag and len(outer) == 2 and outer[0] is inner
        replace = outer if decorated else inner
        parent = replace.getparent()
        level = 0
        child = replace
        for ancestor in replace.iterancestors():
            if ancestor.tag in {"msub", "msup", "msubsup", "munder", "mover", "munderover"} and ancestor[0] is not child:
                level += 1
            child = ancestor
        key = "MX" + uuid4().hex[:12].translate(str.maketrans("0123456789abcdef", "ABCDEFGHIJKLMNOP"))
        brace = Brace(key, top, min(level, 2))
        fraction = etree.Element("mfrac")
        main = etree.SubElement(fraction, "mrow")
        etree.SubElement(main, "mtext").text = key + "A"
        main.append(inner[0])
        small = etree.SubElement(fraction, "mrow")
        etree.SubElement(small, "mtext").text = key + "B"
        if decorated:
            small.append(outer[1])
        etree.SubElement(small, "mtext").text = key + "C"
        parent.replace(replace, fraction)
        braces.append(brace)
    return braces


# MTEF 5 TMPL record: options, selector, variation, options. Fraction and
# horizontal brace both start with two LINE slots. The latter adds a stretchy
# CHAR in the fnEXPAND typeface. Validate our marked slots before rewriting.
FRACTION = bytes((3, 0, 11, 0, 0))
LINE_START = re.compile(rb"(?:\x0f[\x00-\x7f])?\x01\x00(?:\x0f[\x00-\x7f])?$")


def restore_native(data, braces):
    if data[:1] != b"\x05":
        raise ValueError("水平大括号需要 MathType MTEF 5。")
    for brace in braces:
        markers = [b"".join(b"\x02\x00\x81" + ch.encode("utf-16le")
                            for ch in brace.key + suffix) for suffix in "ABC"]
        if any(data.count(marker) != 1 for marker in markers):
            raise ValueError("MathType 未保留用于构建大括号的公式槽。")
        a, b, c = (data.index(marker) for marker in markers)
        length = len(markers[0])
        start = data.rfind(FRACTION, max(0, a - 32), a)
        second = LINE_START.search(data, max(a + length, b - 10), b)
        if (not a < b < c or start < 0 or not second
                or not LINE_START.search(data, start + len(FRACTION), a)
                or data[c+length:c+length+1] != b"\0"):
            raise ValueError("MathType 返回了无法识别的大括号槽结构。")
        base_size = 10 + brace.level
        small_size = min(12, base_size + 1)
        header = bytes((3, 0, 24, int(brace.top), 0))
        fence = bytes((0, base_size, 2, 0, 128 + 22)) + ("︷" if brace.top else "︸").encode("utf-16le")
        data = b"".join((data[:start], header, data[start+5:a],
                         data[a+length:second.start()], bytes((small_size,)),
                         data[second.start():b], data[b+length:c], fence, data[c+length+1:]))
    return data
