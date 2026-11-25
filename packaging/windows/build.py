"""Build the self-contained Windows x64 installer from a clean, relocatable runtime."""

import argparse
from collections import defaultdict
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
from urllib.request import urlopen
from zipfile import ZipFile


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
CACHE = ROOT / ".work" / "pack"
RELEASE = ROOT / "release"
VERSION = "1.0.0"
PYTHON = "3.12.13"
NSIS_URL = "https://downloads.sourceforge.net/project/nsis/NSIS%203/3.11/nsis-3.11.zip"
NSIS_HASH = "c7d27f780ddb6cffb4730138cd1591e841f4b7edb155856901cdf5f214394fa1"
MODEL_URL = "https://www.modelscope.cn/models/RapidAI/"
MODELS = [
    ("rapid_doc/model/formula/rapid_formula_self/models/pp_formulanet_plus_m.onnx",
     "RapidDoc/resolve/v1.0.1/formula/PP-FormulaNet_plus-M/pp_formulanet_plus_m.onnx",
     "52afa60ff51ad66cd7840a74188a058fe99081e617d3bd780117a2d67d94f50d"),
    ("rapid_doc/model/layout/rapid_layout_self/models/pp_doclayoutv3.onnx",
     "RapidDoc/resolve/master/layout/PP-DocLayoutV3/pp_doclayoutv3.onnx",
     "b9a2759e51ee2cc9d98f10cea21caf5862043b882b5f57c683eaf6fa247196a3"),
    ("rapidocr/models/ch_ppocr_mobile_v2.0_cls_mobile.onnx",
     "RapidOCR/resolve/v3.10.0/onnx/PP-OCRv4/cls/ch_ppocr_mobile_v2.0_cls_mobile.onnx",
     "e47acedf663230f8863ff1ab0e64dd2d82b838fceb5957146dab185a89d6215c"),
    ("rapidocr/models/PP-OCRv6_det_small.onnx",
     "RapidOCR/resolve/v3.10.0/onnx/PP-OCRv6/det/PP-OCRv6_det_small.onnx",
     "090f04abcd9d9a7498bc4ebf677e4cb9bdce1fe4197ddb7e529f1ef44e1ff94f"),
    ("rapidocr/models/PP-OCRv6_rec_small.onnx",
     "RapidOCR/resolve/v3.10.0/onnx/PP-OCRv6/rec/PP-OCRv6_rec_small.onnx",
     "6f327246b50388f3c176ae304bd95767ea6dc0c9ae92153ef8cbe210b3c14884"),
]


def run(*args):
    subprocess.run([str(arg) for arg in args], check=True)


def digest(path):
    with path.open("rb") as handle:
        return hashlib.file_digest(handle, "sha256").hexdigest()


def download(path, url, expected):
    if path.is_file() and digest(path) == expected:
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_suffix(".part")
    print(f"Downloading {path.name}", flush=True)
    with urlopen(url, timeout=120) as source, partial.open("wb") as output:
        shutil.copyfileobj(source, output)
    if digest(partial) != expected:
        raise ValueError(f"SHA-256 mismatch for {path.name}")
    partial.replace(path)
    return path


def compiler():
    archive = download(CACHE / "nsis-3.11.zip", NSIS_URL, NSIS_HASH)
    tools = CACHE / f"nsis-{NSIS_HASH[:12]}"
    executable = tools / "makensis.exe"
    if not executable.is_file():
        with TemporaryDirectory(prefix="nsis-", dir=CACHE) as temp, ZipFile(archive) as package:
            package.extractall(temp)
            try:
                (Path(temp) / "nsis-3.11").rename(tools)
            except FileExistsError:
                if not executable.is_file():
                    raise
    return executable


def crt_directory():
    vswhere = Path(os.environ["ProgramFiles(x86)"]) / "Microsoft Visual Studio/Installer/vswhere.exe"
    raw = subprocess.check_output([str(vswhere), "-all", "-products", "*", "-utf8", "-format", "json"])
    folders = [folder for vs in json.loads(raw.decode("utf-8-sig"))
               for folder in Path(vs["installationPath"]).glob("VC/Redist/MSVC/*/x64/Microsoft.*.CRT")]
    if not folders:
        raise RuntimeError("Install Visual Studio Build Tools with the x64 C++ redistributable.")
    return max(folders, key=lambda folder: tuple(int(n) for n in folder.parents[1].name.split(".")))


def prepare(payload, model_cache):
    key = f"cpython-{PYTHON}-windows-x86_64-none"
    run("uv", "python", "install", key, "--install-dir", CACHE / "py", "--no-bin", "--no-registry")
    runtime = payload / "runtime"
    shutil.copytree(CACHE / "py" / key, runtime)
    run("uv", "pip", "install", "--python", runtime / "python.exe", "--link-mode", "copy",
        "--break-system-packages", "--no-config", "-r", HERE / "requirements.lock")
    site = runtime / "Lib/site-packages"
    shutil.copytree(ROOT / "tex2word", site / "tex2word", ignore=shutil.ignore_patterns("__pycache__"))
    for rel, url, expected in MODELS:
        local = model_cache / rel if model_cache else CACHE / "models" / rel
        if not local.is_file() or digest(local) != expected:
            local = download(CACHE / "models" / rel, MODEL_URL + url, expected)
        dest = site / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(local, dest)
    for dll in crt_directory().glob("*.dll"):
        if not (runtime / dll.name).exists():
            shutil.copy2(dll, runtime / dll.name)
    csc = Path(os.environ["WINDIR"]) / "Microsoft.NET/Framework64/v4.0.30319/csc.exe"
    wpf = csc.parent / "WPF"
    desktop = ROOT / "desktop"
    run(csc, "/nologo", "/warnaserror+", "/target:winexe", "/platform:x64", "/codepage:65001", "/optimize+",
        f"/r:{wpf / 'PresentationFramework.dll'}", f"/r:{wpf / 'PresentationCore.dll'}",
        f"/r:{wpf / 'WindowsBase.dll'}", "/r:System.Xaml.dll",
        f"/resource:{desktop / 'MainWindow.xaml'},TeX2Word.MainWindow.xaml",
        f"/out:{payload / 'TeX2Word.exe'}",
        f"/win32icon:{desktop / 'TeX2Word.ico'}", f"/win32manifest:{desktop / 'app.manifest'}",
        desktop / "Program.cs", desktop / "Conversion.cs", desktop / "CliRunner.cs",
        desktop / "ShellDialogs.cs", desktop / "Theme.cs")
    run(csc, "/nologo", "/warnaserror+", "/target:exe", "/platform:x64", "/codepage:65001", "/optimize+",
        f"/out:{payload / 'TeX2Word.com'}", f"/win32icon:{desktop / 'TeX2Word.ico'}",
        desktop / "Cli.cs", desktop / "CliRunner.cs")
    shutil.copy2(HERE / "third-party.txt", payload)


def file_lists(payload, work):
    # The installer and uninstaller own precisely the same files. A user-created
    # document in the installation directory survives uninstall.
    excluded = ("runtime/include/", "runtime/libs/", "runtime/Scripts/", "runtime/Lib/test/",
                "runtime/Lib/ensurepip/", "runtime/Lib/site-packages/pip/",
                "runtime/Lib/site-packages/onnx/backend/test/data/")
    grouped = defaultdict(list)
    size = 0
    for path in sorted(payload.rglob("*")):
        rel = path.relative_to(payload)
        if (not path.is_file() or rel.as_posix().startswith(excluded) or "__pycache__" in rel.parts
                or path.suffix in {".pyc", ".pyo"}):
            continue
        grouped[rel.parent].append(path)
        size += path.stat().st_size

    def escape(value):
        return str(value).replace("$", "$$")

    install, remove, dirs = [], [], set()
    for folder, files in grouped.items():
        dest = "$INSTDIR" if folder == Path(".") else "$INSTDIR\\" + escape(folder)
        install.append(f'SetOutPath "{dest}"')
        for path in files:
            install.append(f'File "{escape(path)}"')
            remove.append(f'Delete "{dest}\\{escape(path.name)}"')
        dirs.update((folder, *folder.parents))
    for folder in sorted(dirs, key=lambda p: (-len(p.parts), str(p))):
        if folder != Path("."):
            remove.append(f'RMDir "$INSTDIR\\{escape(folder)}"')
    for name, lines in (("install.nsh", install), ("remove.nsh", remove)):
        (work / name).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return (size + 1023) // 1024


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--model-cache", type=Path, help="Optional site-packages with previously downloaded OCR models")
    args = parser.parse_args()
    CACHE.mkdir(parents=True, exist_ok=True)
    RELEASE.mkdir(exist_ok=True)
    makensis = compiler()
    with TemporaryDirectory(prefix="build-", dir=CACHE) as folder:
        work = Path(folder)
        payload = work / "app"
        payload.mkdir()
        prepare(payload, args.model_cache)
        size = file_lists(payload, work)
        name = f"TeX2Word-Setup-{VERSION}-x64.exe"
        print("Compressing the Windows installer…", flush=True)
        run(makensis, "/V2", "/INPUTCHARSET", "UTF8", f"/DVERSION={VERSION}", f"/DSIZEKB={size}",
            f"/DICONFILE={ROOT / 'desktop/TeX2Word.ico'}", f"/DOUTPUT={work / name}",
            f"/DINSTALLFILES={work / 'install.nsh'}", f"/DREMOVEFILES={work / 'remove.nsh'}", HERE / "setup.nsi")
        installer = (work / name).replace(RELEASE / name)
        print(f"{installer}\n{installer.stat().st_size:,} bytes\nSHA-256 {digest(installer)}", flush=True)


if __name__ == "__main__":
    main()
