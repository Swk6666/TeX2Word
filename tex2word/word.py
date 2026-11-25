"""Create and verify native MathType OLE objects using the installed desktop apps."""

import ctypes
from contextlib import contextmanager
from io import BytesIO
import os
from pathlib import Path
import posixpath
import re
import struct
import winreg
from zipfile import ZipFile

import olefile
import pythoncom
import win32com.client
import win32event
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from lxml import etree

from .braces import restore_native


def style_document(path, manuscript):
    doc = Document(path)
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.left_margin = section.right_margin = Mm(25)
        section.top_margin, section.bottom_margin = Mm(23), Mm(22)
        if manuscript.has_toc:
            section.different_first_page_header_footer = True
            footer = section.footer.paragraphs[0]
            footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
            page_field = OxmlElement("w:fldSimple")
            page_field.set(qn("w:instr"), "PAGE")
            footer._p.append(page_field)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.20
    for name in ["Normal", "Body Text", "First Paragraph", "Title", "Date", "TOC Heading",
                 *(f"Heading {n}" for n in range(1, 10))]:
        if name not in doc.styles:
            continue
        style = doc.styles[name]
        is_heading = name == "Title" or "Heading" in name
        style.font.name = "Times New Roman"
        fonts = style.element.get_or_add_rPr().rFonts
        for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme"):
            fonts.attrib.pop(qn(f"w:{key}"), None)
        fonts.set(qn("w:eastAsia"), "黑体" if is_heading else "宋体")
        style.font.color.rgb = RGBColor(0, 0, 0)
        if is_heading:
            style.font.bold = True
        if name == "Title":
            style.font.size = Pt(24)
            style.paragraph_format.space_before = Pt(105 if manuscript.has_toc else 14)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if name == "Date":
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if name == "Heading 1":
            style.font.size = Pt(17)
            style.paragraph_format.page_break_before = manuscript.book and manuscript.has_toc
        if name == "Heading 2":
            style.font.size = Pt(14)
        if name == "TOC Heading":
            style.font.size = Pt(19)
    doc.save(path)


def sdk_path():
    arch = str(struct.calcsize("P") * 8)
    roots = []
    if override := os.environ.get("MATHTYPE_DLL"):
        if not Path(override).is_file():
            raise FileNotFoundError(f"MATHTYPE_DLL 文件不存在：{override}")
        return Path(override)
    try:
        cls = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, r"Equation.DSMT4\CLSID")
        command = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{cls}\LocalServer32")
        exe = re.search(r'(?i)"?([^"]+?\.exe)\b', command)
        if exe:
            roots.append(Path(exe[1]).parent)
    except OSError:
        pass
    for key in ("ProgramFiles(x86)", "ProgramFiles"):
        if os.environ.get(key):
            roots.append(Path(os.environ[key]) / "MathType")
    for root in roots:
        for path in (root / "System" / arch / "MT6.dll", root / "System" / "MT6.dll"):
            if path.is_file():
                return path
    raise FileNotFoundError("找不到桌面版 MathType 的 MT6.dll。自定义路径可设置 MATHTYPE_DLL。")


def check(status, operation):
    if status:
        raise RuntimeError(f"MathType {operation} 失败，状态码 {status}")


@contextmanager
def mathtype():
    sdk = ctypes.WinDLL(str(sdk_path()))
    ptr, short, long = ctypes.c_void_p, ctypes.c_short, ctypes.c_long
    for name, args in {
        "MTAPIConnect": [short, short], "MTAPIDisconnect": [],
        "MTSetEqnFromLangStr": [ptr, short, ptr, long],
        "MTGetLangStrFromEqn": [ptr, short, ptr, ctypes.POINTER(long)],
        "MTGetLangBytesFromEqn": [ptr, short, ptr, ctypes.POINTER(long)],
        "MTSetEqnFromLangBytes": [ptr, short, ptr, long],
        "MTCloseOleObject": [long, ptr],
    }.items():
        fn = getattr(sdk, name)
        fn.argtypes, fn.restype = args, long
    mutex = win32event.CreateMutex(None, False, r"Local\TeX2WordMathType")
    win32event.WaitForSingleObject(mutex, win32event.INFINITE)
    try:
        check(sdk.MTAPIConnect(0, 30), "连接")
        try:
            yield sdk
        finally:
            sdk.MTAPIDisconnect()
    finally:
        win32event.ReleaseMutex(mutex)
        mutex.Close()


@contextmanager
def editing(shape, sdk):
    if shape.OLEFormat.ProgID != "Equation.DSMT4":
        raise ValueError("Word 未创建 Equation.DSMT4 对象。")
    shape.OLEFormat.DoVerb(2)
    dispatch = shape.OLEFormat.Object
    # pywin32 keeps the dispatch alive and exposes its native address in repr.
    # Keep 'dispatch' in scope until MTCloseOleObject has returned.
    address = re.search(r"with obj at (0x[0-9a-fA-F]+)", repr(dispatch._oleobj_))
    if not address:
        raise RuntimeError("无法获取 MathType 的 IDispatch 接口。")
    ptr = int(address[1], 16)
    try:
        yield ptr
    except BaseException:
        sdk.MTCloseOleObject(0, ptr)
        raise
    else:
        check(sdk.MTCloseOleObject(1, ptr), "关闭公式")


def insert_mathml(sdk, ptr, xml):
    """Insert at the MathType caret. Newly created equation objects start empty."""
    buf = ctypes.create_unicode_buffer(xml)
    check(sdk.MTSetEqnFromLangStr(ptr, 2, buf, len(xml.encode("utf-16-le")) // 2), "写入公式")


def read_mathml(sdk, ptr):
    size = ctypes.c_long(1 << 20)
    buf = ctypes.create_string_buffer(size.value)
    check(sdk.MTGetLangStrFromEqn(ptr, 2, buf, ctypes.byref(size)), "回读公式")
    # MathType may include translator comments around the root element.
    root = etree.fromstring(buf.value)
    if etree.QName(root).localname != "math" or not "".join(root.itertext()).strip():
        raise ValueError("MathType 回读了空公式。")
    return root


def brace_mtef(doc, sdk, eq):
    scratch = doc.Application.Documents.Add(Visible=False)
    try:
        shape = scratch.Content.InlineShapes.AddOLEObject(ClassType="Equation.DSMT4")
        with editing(shape, sdk) as ptr:
            insert_mathml(sdk, ptr, eq.mathml)
            size = ctypes.c_long()
            check(sdk.MTGetLangBytesFromEqn(ptr, 3, None, ctypes.byref(size)), "读取 MTEF 大小")
            buffer = ctypes.create_string_buffer(size.value)
            check(sdk.MTGetLangBytesFromEqn(ptr, 3, buffer, ctypes.byref(size)), "读取 MTEF")
        return restore_native(buffer.raw, eq.braces)
    finally:
        scratch.Close(0)


@contextmanager
def word_document(path, read_only=False):
    pythoncom.CoInitialize()
    app = doc = None
    try:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3
        doc = app.Documents.Open(str(path.resolve()), ReadOnly=read_only, AddToRecentFiles=False,
                                 ConfirmConversions=False, Visible=False)
        yield doc
    finally:
        try:
            if doc is not None:
                doc.Close(0)
        finally:
            try:
                if app is not None:
                    app.Quit(0)
            finally:
                pythoncom.CoUninitialize()


def find_marker(doc, marker):
    # Main text, tables, notes, headers and footers use distinct Word stories.
    for story in doc.StoryRanges:
        while story is not None:
            where = story.Duplicate
            search = where.Find
            search.ClearFormatting()
            if search.Execute(FindText=marker, Forward=True, Wrap=0, Format=False,
                              MatchWildcards=False, MatchWholeWord=False, MatchCase=True):
                return where
            story = story.NextStoryRange
    raise ValueError(f"Word 中缺少公式占位符 {marker}。")


def embed(path, equations, progress):
    with word_document(path) as doc:
        if equations:
            with mathtype() as sdk:
                for i, eq in enumerate(equations, 1):
                    native = brace_mtef(doc, sdk, eq) if eq.braces else None
                    anchor = find_marker(doc, eq.marker)
                    anchor.Text = ""
                    shape = anchor.InlineShapes.AddOLEObject(ClassType="Equation.DSMT4", Range=anchor)
                    try:
                        with editing(shape, sdk) as ptr:
                            if native is None:
                                insert_mathml(sdk, ptr, eq.mathml)
                            else:
                                check(sdk.MTSetEqnFromLangBytes(ptr, 3, native, len(native)), "写入大括号")
                            result = read_mathml(sdk, ptr)
                            if "\ufffd" in "".join(result.itertext()):
                                raise ValueError("存在 MathType 无法表示的字符")
                    except Exception as exc:
                        raise ValueError(f"第 {i} 个公式转换失败（{eq.tex[:85]}）：{exc}") from exc
                    limit = 453.5 - (72 if eq.number else 0)
                    if shape.Width > limit:
                        shape.LockAspectRatio = -1
                        shape.Width = limit
                    shape.AlternativeText = eq.tex
                    if i == len(equations) or i == 1 or i % 10 == 0:
                        progress(f"MathType 已写入并回读 {i}/{len(equations)} 个公式")
        doc.Fields.Update()
        doc.Repaginate()
        for toc in doc.TablesOfContents:
            toc.Update()
            doc.Repaginate()
            toc.UpdatePageNumbers()
        doc.Save()


def verify(path, expected=None):
    """Follow each embedded OLE relationship and inspect the Equation Native stream."""
    ns = {"o": "urn:schemas-microsoft-com:office:office",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
    count = 0
    with ZipFile(path) as z:
        for part in z.namelist():
            if not re.fullmatch(r"word/(?:document|footnotes|endnotes|header\d+|footer\d+)\.xml", part):
                continue
            root = etree.fromstring(z.read(part))
            if root.xpath("//m:oMath", namespaces=ns):
                raise ValueError(f"{part} 含未转换的 Word 原生公式。")
            objects = root.xpath('//o:OLEObject[@ProgID="Equation.DSMT4"]', namespaces=ns)
            if not objects:
                continue
            folder, name = posixpath.split(part)
            rels = etree.fromstring(z.read(f"{folder}/_rels/{name}.rels"))
            targets = {r.get("Id"): r.get("Target") for r in rels
                       if r.get("Type", "").endswith("/oleObject") and r.get("TargetMode") != "External"}
            for obj in objects:
                rid = obj.get("{" + ns["r"] + "}id")
                if obj.get("Type") != "Embed" or rid not in targets:
                    raise ValueError("公式缺少内嵌 OLE 数据。")
                target = posixpath.normpath(posixpath.join(folder, targets[rid])).lstrip("/")
                with olefile.OleFileIO(BytesIO(z.read(target))) as ole:
                    native = ole.openstream("Equation Native").read()
                    if len(native) < 34:
                        raise ValueError("MathType Equation Native 流过短。")
                    header, = struct.unpack_from("<H", native)
                    length, = struct.unpack_from("<I", native, 8)
                    if not (header >= 12 and length > 5 and header + length <= len(native)
                            and native[header] in (3, 4, 5)):
                        raise ValueError("MathType MTEF 数据不完整。")
                count += 1
    if expected is not None and count != expected:
        raise ValueError(f"需要 {expected} 个 MathType 对象，实际读到 {count} 个。")
    return count
