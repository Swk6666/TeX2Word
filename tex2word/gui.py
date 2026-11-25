import os
from pathlib import Path
from queue import Empty, SimpleQueue
import subprocess
import sys
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def run():
    root = tk.Tk()
    root.title("TeX2Word · MathType 转换")
    root.minsize(620, 288)
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    source, target = tk.StringVar(), tk.StringVar()
    language, status = tk.StringVar(value="ch"), tk.StringVar(value="请选择 TeX 主文件或 PDF。")
    events = SimpleQueue()
    busy = closing = False

    def choose_input():
        path = filedialog.askopenfilename(filetypes=[("TeX 与 PDF", "*.tex *.pdf"), ("所有文件", "*.*")])
        if path:
            p = Path(path)
            source.set(path)
            target.set(str(p.with_name(f"{p.stem}-{p.suffix[1:].lower()}-MathType.docx")))

    def choose_output():
        current = Path(target.get() or "MathType.docx")
        path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=current.name,
                                            initialdir=str(current.parent), filetypes=[("Word 文档", "*.docx")])
        if path:
            target.set(path)

    for row, (title, variable, command) in enumerate([
        ("输入文件", source, choose_input), ("Word 文档", target, choose_output)
    ]):
        ttk.Label(frame, text=title).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(frame, textvariable=variable, width=48).grid(row=row, column=1, sticky="ew", pady=6)
        ttk.Button(frame, text="浏览…", command=command).grid(row=row, column=2, padx=(8, 0))
    ttk.Label(frame, text="PDF 语言").grid(row=2, column=0, sticky="w", pady=7)
    languages = ttk.Frame(frame)
    languages.grid(row=2, column=1, sticky="w")
    for value, label in [("ch", "中文或中英混排"), ("en", "英文")]:
        ttk.Radiobutton(languages, text=label, variable=language, value=value).pack(side="left", padx=(0, 12))

    bar = ttk.Progressbar(frame, mode="indeterminate")
    bar.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(12, 7))
    ttk.Label(frame, textvariable=status, wraplength=560).grid(row=4, column=0, columnspan=3, sticky="w")
    actions = ttk.Frame(frame)
    actions.grid(row=5, column=0, columnspan=3, sticky="e", pady=(16, 0))

    def worker(args, destination):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        last = ""
        try:
            with subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                  text=True, encoding="utf-8", errors="replace", env=env,
                                  cwd=Path(__file__).resolve().parents[1],
                                  creationflags=subprocess.CREATE_NO_WINDOW) as proc:
                for line in proc.stdout:
                    line = line.strip()
                    if line:
                        last = line
                    if line.startswith(("正在", "识别到", "已识别", "MathType 已", "未解析", "未排版", "已保存")):
                        events.put(("progress", line))
                code = proc.wait()
            if code:
                events.put(("error", last or f"退出码 {code}"))
            else:
                events.put(("done", destination))
        except OSError as exc:
            events.put(("error", str(exc)))

    def poll():
        nonlocal busy
        try:
            while True:
                kind, value = events.get_nowait()
                if kind == "progress":
                    status.set(value)
                else:
                    busy = False
                    bar.stop()
                    start.configure(state="normal")
                    if closing:
                        root.destroy()
                        return
                    if kind == "error":
                        status.set("转换失败")
                        messagebox.showerror("转换失败", value, parent=root)
                    else:
                        open_button.configure(state="normal", command=lambda p=value: os.startfile(p))
        except Empty:
            pass
        if busy:
            root.after(150, poll)

    def begin():
        nonlocal busy
        p, dest = Path(source.get().strip()), Path(target.get().strip())
        if not p.is_file() or p.suffix.lower() not in {".tex", ".pdf"}:
            messagebox.showerror("输入文件", "请选择已有的 .tex 或 .pdf 文件。", parent=root)
            return
        if dest.suffix.lower() != ".docx":
            messagebox.showerror("输出路径", "请选择 .docx 输出文件。", parent=root)
            return
        force = dest.exists()
        if force and not messagebox.askyesno("替换 Word 文档", f"替换已有文件？\n{dest}", parent=root):
            return
        exe = Path(sys.executable).with_name("python.exe")
        args = [str(exe), "-u", "-m", "tex2word", str(p.resolve()), "-o", str(dest.resolve()),
                "--lang", language.get()]
        if force:
            args.append("--force")
        busy = True
        status.set("正在读取文档…")
        start.configure(state="disabled")
        open_button.configure(state="disabled")
        bar.start(14)
        Thread(target=worker, args=(args, str(dest.resolve())), daemon=True).start()
        root.after(150, poll)

    def close():
        nonlocal closing
        if busy:
            closing = True
            root.withdraw()
        else:
            root.destroy()

    start = ttk.Button(actions, text="开始转换", command=begin)
    start.pack(side="left", padx=(0, 8))
    open_button = ttk.Button(actions, text="打开结果", state="disabled")
    open_button.pack(side="left")
    root.protocol("WM_DELETE_WINDOW", close)
    root.mainloop()
