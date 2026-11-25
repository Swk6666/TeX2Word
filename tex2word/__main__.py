import argparse
from pathlib import Path
import sys

from . import __version__


def main(argv=None):
    parser = argparse.ArgumentParser(
        prog="tex2word",
        description="将 PDF、TeX 或 Markdown 转为含 MathType 可编辑公式的 Word 文档",
    )
    parser.add_argument("sources", nargs="*", type=Path, help="输入 .pdf、.tex 或 .md 文件（支持多个）")
    parser.add_argument("-o", "--output", type=Path, help="输出 .docx 路径（仅单文件转换时使用）")
    parser.add_argument("-d", "--output-dir", type=Path, help="批量转换时统一输出到此目录")
    parser.add_argument("--lang", choices=["ch", "en"], default="ch", help="PDF 识别语言（默认中英混排）")
    parser.add_argument("--date", help=r"指定 TeX 标题页日期，例如原 PDF 编译 \today 时的日期")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    parser.add_argument("--verify", type=Path, metavar="DOCX", help="检查现有文档的 MathType 原生数据")
    parser.add_argument("-V", "--version", action="version", version=f"TeX2Word {__version__}")
    args = parser.parse_args(argv)
    try:
        if args.verify:
            from .word import verify
            print(f"已验证 {verify(args.verify)} 个 MathType 原生对象", flush=True)
            return 0
        if not args.sources:
            from .gui import run
            run()
            return 0
        if args.output and len(args.sources) > 1:
            parser.error("指定多个输入文件时请使用 -d/--output-dir，而非 -o/--output。")
        if args.output and args.output_dir:
            parser.error("-o/--output 与 -d/--output-dir 不能同时使用。")
        from .convert import convert
        failed = 0
        for index, source in enumerate(args.sources, 1):
            target = args.output
            if args.output_dir:
                ext = source.suffix.lower().lstrip(".")
                target = args.output_dir / f"{source.stem}-{ext}-MathType.docx"
            prefix = f"[{index}/{len(args.sources)}] " if len(args.sources) > 1 else ""
            try:
                path, count = convert(
                    source, target, lang=args.lang, date=args.date, overwrite=args.force,
                    progress=lambda msg, p=prefix: print(f"{p}{msg}", flush=True),
                )
                print(f"{prefix}已保存 {path}（{count} 个 MathType 对象）", flush=True)
            except Exception as exc:
                failed += 1
                print(f"{prefix}转换失败（{source}）：{exc}", file=sys.stderr, flush=True)
        return 1 if failed else 0
    except Exception as exc:
        print(f"转换失败：{exc}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
