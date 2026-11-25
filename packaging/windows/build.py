"""Build the Windows x64 installer (lightweight standard or full offline) from a clean runtime."""

import argparse
from collections import defaultdict
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from urllib.request import urlopen
from zipfile import ZipFile

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tex2word.models import MODEL_BASE_URL, OCR_CLS_MODEL, PDF_MODELS

CACHE = ROOT / ".work" / "pack"
RELEASE = ROOT / "release"
VERSION = "1.0.0"
PYTHON = "3.12.13"
NSIS_URL = "https://downloads.sourceforge.net/project/nsis/NSIS%203/3.11/nsis-3.11.zip"
NSIS_HASH = "c7d27f780ddb6cffb4730138cd1591e841f4b7edb155856901cdf5f214394fa1"

# Directories / files never used by the installed WPF + ONNX CPU Windows app.
PRUNE_DIRS = (
    "runtime/include",
    "runtime/libs",
    "runtime/Scripts",
    "runtime/tcl",
    "runtime/Lib/test",
    "runtime/Lib/ensurepip",
    "runtime/Lib/tkinter",
    "runtime/Lib/idlelib",
    "runtime/Lib/turtledemo",
    "runtime/Lib/site-packages/pip",
    "runtime/Lib/site-packages/pythonwin",
    "runtime/Lib/site-packages/win32/Demos",
    "runtime/Lib/site-packages/openai",
    "runtime/Lib/site-packages/hf_xet",
    "runtime/Lib/site-packages/onnx/backend/test/data",
)

PRUNE_GLOBS = (
    "runtime/DLLs/tcl*.dll",
    "runtime/DLLs/tk*.dll",
    "runtime/DLLs/_tkinter.pyd",
    "runtime/Lib/site-packages/cv2/opencv_videoio_ffmpeg*.dll",
    "runtime/Lib/site-packages/rapid_doc/resources/*.safetensors",
    "runtime/Lib/site-packages/rapid_doc/resources/pp-ocrv4_mobile_seal_det.onnx",
    "runtime/Lib/site-packages/rapidocr/models/PP-OCRv6_det_small.onnx",
    "runtime/Lib/site-packages/rapidocr/models/PP-OCRv6_rec_small.onnx",
    "runtime/Lib/site-packages/PyWin32.chm",
)


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


def prune_payload(payload: Path) -> None:
    """Remove unused GUI/video/PyTorch-only/duplicate files to keep the installer lean."""
    for rel in PRUNE_DIRS:
        target = payload / rel
        if target.is_dir():
            shutil.rmtree(target)
    for pattern in PRUNE_GLOBS:
        for path in payload.glob(pattern):
            if path.is_file():
                path.unlink()
            elif path.is_dir():
                shutil.rmtree(path)


def prepare(payload, model_cache, offline: bool):
    key = f"cpython-{PYTHON}-windows-x86_64-none"
    run("uv", "python", "install", key, "--install-dir", CACHE / "py", "--no-bin", "--no-registry")
    runtime = payload / "runtime"
    shutil.copytree(CACHE / "py" / key, runtime)
    run("uv", "pip", "install", "--python", runtime / "python.exe", "--link-mode", "copy",
        "--break-system-packages", "--no-config", "-r", HERE / "requirements.lock")
    site = runtime / "Lib/site-packages"
    shutil.copytree(ROOT / "tex2word", site / "tex2word", ignore=shutil.ignore_patterns("__pycache__"))

    # Always bundle the tiny 0.56 MB OCR orientation classifier; only bundle the two
    # large (~690 MB) layout + formula ONNX models when --offline is requested.
    bundled_models = (OCR_CLS_MODEL, *PDF_MODELS) if offline else (OCR_CLS_MODEL,)
    for spec in bundled_models:
        local = model_cache / spec.bundled_rel if model_cache else CACHE / "models" / spec.bundled_rel
        if not local.is_file() or digest(local) != spec.sha256:
            local = download(CACHE / "models" / spec.bundled_rel, MODEL_BASE_URL + spec.url_path, spec.sha256)
        dest = site / spec.bundled_rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(local, dest)

    for dll in crt_directory().glob("*.dll"):
        if not (runtime / dll.name).exists():
            shutil.copy2(dll, runtime / dll.name)

    prune_payload(payload)

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
    grouped = defaultdict(list)
    size = 0
    for path in sorted(payload.rglob("*")):
        rel = path.relative_to(payload)
        if not path.is_file() or "__pycache__" in rel.parts or path.suffix in {".pyc", ".pyo"}:
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
    parser.add_argument("--offline", action="store_true",
                        help="Bundle the ~690 MB PDF layout/formula ONNX models for air-gapped offline installation")
    args = parser.parse_args()
    CACHE.mkdir(parents=True, exist_ok=True)
    RELEASE.mkdir(exist_ok=True)
    makensis = compiler()
    with TemporaryDirectory(prefix="build-", dir=CACHE) as folder:
        work = Path(folder)
        payload = work / "app"
        payload.mkdir()
        prepare(payload, args.model_cache, offline=args.offline)
        size = file_lists(payload, work)
        suffix = "-offline" if args.offline else ""
        name = f"TeX2Word-Setup-{VERSION}-x64{suffix}.exe"
        print(f"Compressing the Windows installer ({name}, LZMA solid)…", flush=True)
        defines = [
            f"/DVERSION={VERSION}", f"/DSIZEKB={size}",
            f"/DICONFILE={ROOT / 'desktop/TeX2Word.ico'}", f"/DOUTPUT={work / name}",
            f"/DINSTALLFILES={work / 'install.nsh'}", f"/DREMOVEFILES={work / 'remove.nsh'}",
        ]
        if args.offline:
            defines.append("/DOFFLINE_MODELS=1")
        run(makensis, "/V2", "/INPUTCHARSET", "UTF8", *defines, HERE / "setup.nsi")
        installer = (work / name).replace(RELEASE / name)
        print(f"{installer}\n{installer.stat().st_size:,} bytes ({installer.stat().st_size / 1048576:.1f} MB)\nSHA-256 {digest(installer)}", flush=True)


if __name__ == "__main__":
    main()
