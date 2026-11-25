"""Create and verify native MathType OLE objects using the installed desktop apps."""

import ctypes
from contextlib import contextmanager
from io import BytesIO
import os
from pathlib import Path
import posixpath
import re
import struct
import time
import winreg
from zipfile import ZipFile

import olefile
import pythoncom
import win32api
import win32com.client
import win32con
import win32event
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor
from lxml import etree

from .braces import restore_native


def _set_style_fonts(style, east_asia: str, latin: str = "Times New Roman"):
    style.font.name = latin
    fonts = style.element.get_or_add_rPr().get_or_add_rFonts()
    for key in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        fonts.attrib.pop(qn(f"w:{key}"), None)
    fonts.set(qn("w:ascii"), latin)
    fonts.set(qn("w:hAnsi"), latin)
    fonts.set(qn("w:eastAsia"), east_asia)


def _get_or_add_pstyle(doc, name: str, base: str = "Normal"):
    if name in doc.styles:
        return doc.styles[name]
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if base in doc.styles:
        style.base_style = doc.styles[base]
    return style


def _set_first_line_chars(style, chars: int = 200, twips: int = 440):
    ppr = style.element.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    ind.set(qn("w:firstLineChars"), str(chars))
    ind.set(qn("w:firstLine"), str(twips))


def _set_cell_borders(cell, **edges):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, spec in edges.items():
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        for attr, val in spec.items():
            node.set(qn(f"w:{attr}"), str(val))


def _set_table_borders(table, **edges):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge, spec in edges.items():
        node = borders.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        for attr, val in spec.items():
            node.set(qn(f"w:{attr}"), str(val))


def _set_table_width(table, width_dxa: int):
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(width_dxa))
    tbl_w.set(qn("w:type"), "dxa")


def _set_cell_margins(table, top=70, bottom=70, left=100, right=100):
    tbl_pr = table._tbl.tblPr
    mar = tbl_pr.find(qn("w:tblCellMar"))
    if mar is None:
        mar = OxmlElement("w:tblCellMar")
        tbl_pr.append(mar)
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            mar.append(node)
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")


def _style_tables(doc, table_kinds):
    none_edge = {"val": "none", "sz": "0", "space": "0", "color": "auto"}
    thick = {"val": "single", "sz": "12", "space": "0", "color": "000000"}
    thin = {"val": "single", "sz": "6", "space": "0", "color": "000000"}

    for idx, table in enumerate(doc.tables):
        kind = table_kinds[idx] if idx < len(table_kinds) else "t2w-table"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True

        if kind == "t2w-table-meta":
            # Cover page key-value table: centered, label bold/right-aligned, value centered with bottom underline.
            _set_table_width(table, 6400)
            _set_table_borders(table, top=none_edge, bottom=none_edge, left=none_edge,
                               right=none_edge, insideH=none_edge, insideV=none_edge)
            _set_cell_margins(table, top=110, bottom=80, left=80, right=80)
            col_widths = (1900, 4500)
            for row in table.rows:
                tr_pr = row._tr.get_or_add_trPr()
                tr_pr.append(OxmlElement("w:cantSplit"))
                for ci, cell in enumerate(row.cells):
                    tc_w = cell._tc.get_or_add_tcPr().find(qn("w:tcW"))
                    if tc_w is None:
                        tc_w = OxmlElement("w:tcW")
                        cell._tc.get_or_add_tcPr().append(tc_w)
                    tc_w.set(qn("w:w"), str(col_widths[min(ci, 1)]))
                    tc_w.set(qn("w:type"), "dxa")
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
                    if ci == 1:
                        _set_cell_borders(cell, bottom=thin)
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if ci == 0 else WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.first_line_indent = Pt(0)
                        for run in p.runs:
                            run.font.size = Pt(14)
                            if ci == 0:
                                run.font.bold = True

        elif kind == "t2w-table-sig":
            # Borderless 2-column signature / date table.
            _set_table_width(table, 8200)
            _set_table_borders(table, top=none_edge, bottom=none_edge, left=none_edge,
                               right=none_edge, insideH=none_edge, insideV=none_edge)
            _set_cell_margins(table, top=80, bottom=80, left=120, right=120)
            for row in table.rows:
                tr_pr = row._tr.get_or_add_trPr()
                tr_pr.append(OxmlElement("w:cantSplit"))
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(3)
                        p.paragraph_format.space_after = Pt(3)
                        p.paragraph_format.first_line_indent = Pt(0)
                        for run in p.runs:
                            run.font.size = Pt(11)

        else:
            # Academic three-line table (三线表).
            _set_table_width(table, 9070)
            _set_table_borders(table, top=thick, bottom=thick, left=none_edge,
                               right=none_edge, insideH=none_edge, insideV=none_edge)
            _set_cell_margins(table, top=65, bottom=65, left=80, right=80)
            for ri, row in enumerate(table.rows):
                tr_pr = row._tr.get_or_add_trPr()
                tr_pr.append(OxmlElement("w:cantSplit"))
                if ri == 0 and len(table.rows) >= 2:
                    tr_pr.append(OxmlElement("w:tblHeader"))
                for cell in row.cells:
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    if ri == 0 and len(table.rows) >= 2:
                        _set_cell_borders(cell, bottom=thin)
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.paragraph_format.space_before = Pt(1)
                        p.paragraph_format.space_after = Pt(1)
                        p.paragraph_format.line_spacing = 1.15
                        p.paragraph_format.first_line_indent = Pt(0)
                        for run in p.runs:
                            run.font.size = Pt(10)


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
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.25

    for name in ["Normal", "Body Text", "First Paragraph", "Compact", "Title", "Date", "TOC Heading",
                 *(f"Heading {n}" for n in range(1, 10))]:
        if name not in doc.styles:
            continue
        style = doc.styles[name]
        is_heading = name == "Title" or "Heading" in name
        _set_style_fonts(style, "黑体" if is_heading else "宋体")
        style.font.color.rgb = RGBColor(0, 0, 0)
        if is_heading:
            style.font.bold = True
            style.paragraph_format.keep_with_next = True
        if name == "Title":
            style.font.size = Pt(24)
            style.paragraph_format.space_before = Pt(105 if manuscript.has_toc else 14)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if name == "Date":
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if name == "Heading 1":
            style.font.size = Pt(16)
            style.paragraph_format.space_before = Pt(6)
            style.paragraph_format.space_after = Pt(12)
            style.paragraph_format.page_break_before = manuscript.book and manuscript.has_toc
        if name == "Heading 2":
            style.font.size = Pt(14)
            style.paragraph_format.space_before = Pt(10)
            style.paragraph_format.space_after = Pt(5)
        if name == "Heading 3":
            style.font.size = Pt(12)
            style.paragraph_format.space_before = Pt(7)
            style.paragraph_format.space_after = Pt(4)
        if name == "TOC Heading":
            style.font.size = Pt(18)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            style.paragraph_format.space_after = Pt(10)

    # Configure semantic custom styles emitted by the PDF layout pipeline.
    custom_specs = {
        "T2WFigure": {"east_asia": "宋体", "size": 11, "align": WD_ALIGN_PARAGRAPH.CENTER,
                      "before": 6, "after": 2, "line": 1.0, "keep_next": True},
        "T2WFigureLeft": {"east_asia": "宋体", "size": 11, "align": WD_ALIGN_PARAGRAPH.LEFT,
                          "before": 0, "after": 8, "line": 1.0, "keep_next": False},
        "T2WFigureCaption": {"east_asia": "宋体", "size": 10, "align": WD_ALIGN_PARAGRAPH.CENTER,
                             "before": 2, "after": 8, "line": 1.15, "keep_together": True},
        "T2WTableCaption": {"east_asia": "宋体", "size": 10.5, "align": WD_ALIGN_PARAGRAPH.CENTER,
                            "before": 8, "after": 3, "line": 1.15, "keep_next": True},
        "T2WBodyIndent": {"east_asia": "宋体", "size": 11, "align": WD_ALIGN_PARAGRAPH.JUSTIFY,
                          "before": 0, "after": 3, "line": 1.30, "first_chars": 200},
        "T2WBodyPlain": {"east_asia": "宋体", "size": 11, "align": WD_ALIGN_PARAGRAPH.JUSTIFY,
                         "before": 0, "after": 4, "line": 1.30},
        "T2WReference": {"east_asia": "宋体", "size": 10.5, "align": WD_ALIGN_PARAGRAPH.JUSTIFY,
                         "before": 0, "after": 3, "line": 1.22, "hang": 22},
        "T2WFrontTitle": {"east_asia": "黑体", "size": 24, "bold": True,
                          "align": WD_ALIGN_PARAGRAPH.CENTER, "before": 48, "after": 22, "line": 1.25},
        "T2WFrontSubtitle": {"east_asia": "黑体", "size": 16, "bold": True,
                             "align": WD_ALIGN_PARAGRAPH.CENTER, "before": 16, "after": 28, "line": 1.30},
        "T2WFrontHeading": {"east_asia": "黑体", "size": 15, "bold": True,
                            "align": WD_ALIGN_PARAGRAPH.CENTER, "before": 16, "after": 10,
                            "line": 1.25, "keep_next": True},
        "T2WFrontCenter": {"east_asia": "宋体", "size": 12, "align": WD_ALIGN_PARAGRAPH.CENTER,
                           "before": 54, "after": 6, "line": 1.25},
    }
    for name, spec in custom_specs.items():
        if name not in doc.styles:
            continue
        st = _get_or_add_pstyle(doc, name)
        _set_style_fonts(st, spec["east_asia"])
        st.font.size = Pt(spec["size"])
        st.font.color.rgb = RGBColor(0, 0, 0)
        if spec.get("bold"):
            st.font.bold = True
        pf = st.paragraph_format
        pf.alignment = spec["align"]
        pf.space_before = Pt(spec["before"])
        pf.space_after = Pt(spec["after"])
        pf.line_spacing = spec["line"]
        pf.keep_with_next = bool(spec.get("keep_next", False))
        pf.keep_together = bool(spec.get("keep_together", False))
        if spec.get("first_chars"):
            _set_first_line_chars(st, spec["first_chars"], int(spec["size"] * 40))
        elif spec.get("hang"):
            pf.left_indent = Pt(spec["hang"])
            pf.first_line_indent = Pt(-spec["hang"])
        else:
            pf.first_line_indent = Pt(0)

    # Apply per-heading centering recorded from PDF geometry without altering left-aligned TeX chapters.
    if manuscript.centered_headings:
        heading_seq = 0
        for p in doc.paragraphs:
            if p.style and p.style.name.startswith("Heading "):
                heading_seq += 1
                if heading_seq in manuscript.centered_headings:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Ensure standalone figure paragraphs without T2WFigure style (e.g. from TeX) are centered cleanly.
    for p in doc.paragraphs:
        if p.style and p.style.name in {"T2WFigure", "T2WFigureLeft"}:
            continue
        if p._p.xpath('.//*[local-name()="drawing"]') and not p.text.strip():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)

    _style_tables(doc, getattr(manuscript, "table_kinds", []))
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


def _embedding_mathtype_pids():
    """Return PIDs of headless MathType -Embedding servers and whether any interactive Word is running."""
    pythoncom.CoInitialize()
    try:
        wmi = win32com.client.GetObject("winmgmts:")
        mt_pids = set()
        interactive_word = False
        for proc in wmi.ExecQuery("SELECT ProcessId, Name, CommandLine FROM Win32_Process "
                                  "WHERE Name='MathType.exe' OR Name='WINWORD.EXE'"):
            name = (proc.Name or "").lower()
            cmd = (proc.CommandLine or "").lower()
            if name == "mathtype.exe" and "-embedding" in cmd:
                mt_pids.add(int(proc.ProcessId))
            elif name == "winword.exe" and "/automation" not in cmd and "-embedding" not in cmd:
                interactive_word = True
        return mt_pids, interactive_word
    except Exception:
        return set(), True
    finally:
        pythoncom.CoUninitialize()


def _terminate_pids(pids):
    for pid in pids:
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
            try:
                win32api.TerminateProcess(handle, 0)
            finally:
                win32api.CloseHandle(handle)
        except Exception:
            pass
    if pids:
        time.sleep(0.3)


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
        # MathType 7's COM server (-Embedding) frequently stays behind after MTAPIDisconnect
        # and deadlocks subsequent Word AddOLEObject calls. Recycle orphaned headless servers
        # when no interactive Word session is open.
        before_pids, interactive_word = _embedding_mathtype_pids()
        if before_pids and not interactive_word:
            _terminate_pids(before_pids)
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
    before_mt, _ = _embedding_mathtype_pids() if equations else (set(), False)
    try:
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
    finally:
        if equations:
            after_mt, interactive_word = _embedding_mathtype_pids()
            if not interactive_word:
                _terminate_pids(after_mt - before_mt)


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
