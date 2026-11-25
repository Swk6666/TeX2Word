import os
from pathlib import Path
from tempfile import TemporaryDirectory

from .manuscript import Manuscript
from .pdf_input import pdf_markdown
from .tex_input import read_tex
from .word import embed, style_document, verify


def convert(source, output=None, *, lang="ch", date=None, overwrite=False, progress=print):
    source = Path(source).resolve(strict=True)
    ext = source.suffix.lower()
    if ext not in {".tex", ".pdf"}:
        raise ValueError("请选择 .tex 或 .pdf 文件。")
    default = source.with_name(f"{source.stem}-{ext[1:]}-MathType.docx")
    target = Path(output or default).resolve()
    if target.suffix.lower() != ".docx":
        raise ValueError("输出文件的扩展名应为 .docx。")
    if target.exists() and not overwrite:
        raise FileExistsError(f"输出已存在：{target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix=".tex2word-", dir=target.parent) as folder:
        work = Path(folder)
        if ext == ".pdf":
            manuscript = Manuscript(pdf_markdown(source, work, lang, progress), "markdown")
        else:
            manuscript = Manuscript(source, "latex", read_tex(source), date)
        staged = work / "result.docx"
        progress("正在整理正文、章节、公式和编号")
        manuscript.write(staged)
        equations = manuscript.equations
        display = sum(eq.display for eq in equations)
        progress(f"识别到 {display} 个独立公式，{len(equations) - display} 个行内公式")
        for warning in sorted(manuscript.warnings):
            progress(warning)
        style_document(staged, manuscript)
        embed(staged, equations, progress)
        count = verify(staged, len(equations))
        # On Windows rename refuses an existing destination, including a racing writer.
        if overwrite:
            os.replace(staged, target)
        else:
            os.rename(staged, target)
    return target, count
