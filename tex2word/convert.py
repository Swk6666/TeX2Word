import os
from pathlib import Path
from shutil import copyfileobj
from tempfile import NamedTemporaryFile, TemporaryDirectory

from .manuscript import Manuscript
from .md_input import read_markdown
from .pdf_input import pdf_markdown
from .tex_input import read_tex
from .word import embed, style_document, verify


def _publish(staged: Path, target: Path, overwrite: bool) -> None:
    """Publish a fully verified docx into target's folder without exposing un-embedded files."""
    with NamedTemporaryFile(prefix=".tex2word-", suffix=".tmp", dir=target.parent, delete=False) as out:
        tmp = Path(out.name)
        try:
            with staged.open("rb") as src:
                copyfileobj(src, out)
        except BaseException:
            out.close()
            tmp.unlink(missing_ok=True)
            raise
    try:
        # On Windows rename refuses an existing destination, including a racing writer.
        if overwrite:
            os.replace(tmp, target)
        else:
            os.rename(tmp, target)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


SUPPORTED_EXTENSIONS = {".tex", ".pdf", ".md", ".markdown"}


def convert(source, output=None, *, lang="ch", date=None, overwrite=False, progress=print):
    source = Path(source).resolve(strict=True)
    ext = source.suffix.lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ValueError("请选择 .pdf、.tex 或 .md 文件。")
    default = source.with_name(f"{source.stem}-{ext[1:]}-MathType.docx")
    target = Path(output or default).resolve()
    if target.suffix.lower() != ".docx":
        raise ValueError("输出文件的扩展名应为 .docx。")
    if target.exists() and not overwrite:
        raise FileExistsError(f"输出已存在：{target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    # Keep intermediate files (which hold MTX placeholders before MathType embedding)
    # in the system temp directory so they can never be mistaken for finished outputs.
    with TemporaryDirectory(prefix="tex2word-") as folder:
        work = Path(folder)
        if ext == ".pdf":
            manuscript = Manuscript(pdf_markdown(source, work, lang, progress), "markdown")
        elif ext in {".md", ".markdown"}:
            manuscript = Manuscript(source, "markdown", read_markdown(source), date)
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
        _publish(staged, target, overwrite)
    return target, count
