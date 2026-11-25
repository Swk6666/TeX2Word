from pathlib import Path
import re


def read_tex(path):
    data = path.read_bytes()
    encoding = re.search(rb"(?im)^%\s*!\s*tex\s+encoding\s*=\s*([A-Za-z0-9_-]+)", data[:2048])
    if encoding:
        return data.decode(encoding[1].decode("ascii")).lstrip("\ufeff")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("gb18030")


# Pandoc's raw_tex reader preserves \input verbatim. Expand literal file references
# before parsing so definitions in included files remain in the main TeX context.
INCLUDE = re.compile(
    r"\\[%\\]|%[^\n]*|\\verb\*?(?P<sep>[^\w\s]).*?(?P=sep)"
    r"|\\begin\{(?P<literal>verbatim\*?|lstlisting|minted|comment)\}.*?\\end\{(?P=literal)\}"
    r"|\\(?P<cmd>input|include)(?![a-zA-Z@])\s*(?:\{(?P<brace>[^{}\\\n]+)\}|(?P<plain>[\w./-]+))",
    re.S,
)


def expand_inputs(path: Path, contents: str):
    def expand(current, text, parents):
        def include(match):
            if not match["cmd"]:
                return match[0]
            name = Path((match["brace"] or match["plain"]).strip().strip('"'))
            candidates = [base / filename for base in (path.parent, current.parent)
                          for filename in (name, Path(str(name) + ".tex"))]
            child = next((p.resolve() for p in candidates if p.is_file()), None)
            if child is None:
                raise FileNotFoundError(f"{current.name} 引用的 TeX 文件不存在：{name}")
            if child in parents:
                raise ValueError(f"TeX 文件循环引用：{child}")
            return "\n" + expand(child, read_tex(child), parents | {child}) + "\n"
        return INCLUDE.sub(include, text)

    return expand(path, contents, {path.resolve()})
