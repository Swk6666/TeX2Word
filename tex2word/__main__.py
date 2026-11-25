import argparse
import sys
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="将 TeX 或 PDF 转为含 MathType 可编辑公式的 Word 文档")
    parser.add_argument("source", nargs="?", type=Path)
    parser.add_argument("-o", "--output", type=Path, help="输出 .docx 路径")
    parser.add_argument("--lang", choices=["ch", "en"], default="ch", help="PDF 识别语言（默认中英混排）")
    parser.add_argument("--date", help=r"指定 TeX 标题页日期，例如原 PDF 编译 \today 时的日期")
    parser.add_argument("--force", action="store_true", help="允许替换已存在的输出文件")
    parser.add_argument("--verify", type=Path, metavar="DOCX", help="检查现有文档的 MathType 原生数据")
    args = parser.parse_args()
    try:
        if args.verify:
            from .word import verify
            print(f"已验证 {verify(args.verify)} 个 MathType 原生对象")
        elif args.source:
            from .convert import convert
            path, count = convert(args.source, args.output, lang=args.lang, date=args.date,
                                  overwrite=args.force, progress=lambda msg: print(msg, flush=True))
            print(f"已保存 {path}（{count} 个 MathType 对象）")
        else:
            from .gui import run
            run()
    except Exception as exc:
        print(f"转换失败：{exc}", file=sys.stderr, flush=True)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
