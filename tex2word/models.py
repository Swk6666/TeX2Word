"""Manage local and on-demand cached ONNX models for offline PDF recognition."""

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import time
from urllib.request import Request, urlopen


MODEL_BASE_URL = "https://www.modelscope.cn/models/RapidAI/"


@dataclass(frozen=True)
class ModelSpec:
    key: str
    filename: str
    label: str
    bundled_rel: str
    url_path: str
    sha256: str


# Small classifier model (~0.56 MB) always bundled with the Windows installer.
OCR_CLS_MODEL = ModelSpec(
    key="ocr_cls",
    filename="ch_ppocr_mobile_v2.0_cls_mobile.onnx",
    label="OCR 方向分类模型",
    bundled_rel="rapidocr/models/ch_ppocr_mobile_v2.0_cls_mobile.onnx",
    url_path="RapidOCR/resolve/v3.10.0/onnx/PP-OCRv4/cls/ch_ppocr_mobile_v2.0_cls_mobile.onnx",
    sha256="e47acedf663230f8863ff1ab0e64dd2d82b838fceb5957146dab185a89d6215c",
)

# Large PDF recognition models (~690 MB total): bundled in --offline builds,
# or downloaded once on demand into %LOCALAPPDATA%\TeX2Word\models in standard builds.
PDF_MODELS: tuple[ModelSpec, ...] = (
    ModelSpec(
        key="layout",
        filename="pp_doclayoutv3.onnx",
        label="PDF 版面分析模型",
        bundled_rel="rapid_doc/model/layout/rapid_layout_self/models/pp_doclayoutv3.onnx",
        url_path="RapidDoc/resolve/master/layout/PP-DocLayoutV3/pp_doclayoutv3.onnx",
        sha256="b9a2759e51ee2cc9d98f10cea21caf5862043b882b5f57c683eaf6fa247196a3",
    ),
    ModelSpec(
        key="formula",
        filename="pp_formulanet_plus_m.onnx",
        label="PDF 公式识别模型",
        bundled_rel="rapid_doc/model/formula/rapid_formula_self/models/pp_formulanet_plus_m.onnx",
        url_path="RapidDoc/resolve/v1.0.1/formula/PP-FormulaNet_plus-M/pp_formulanet_plus_m.onnx",
        sha256="52afa60ff51ad66cd7840a74188a058fe99081e617d3bd780117a2d67d94f50d",
    ),
)


def model_cache_dir() -> Path:
    override = os.environ.get("TEX2WORD_MODEL_DIR")
    if override:
        return Path(override).expanduser().resolve()
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "TeX2Word" / "models"
    return Path.home() / ".cache" / "tex2word" / "models"


def _bundled_site_packages() -> Path | None:
    try:
        import rapid_doc
        return Path(rapid_doc.__file__).resolve().parents[1]
    except Exception:
        return None


def _is_valid_cached(path: Path) -> bool:
    try:
        return path.is_file() and path.stat().st_size > 1_000_000
    except OSError:
        return False


def _download_model(spec: ModelSpec, dest: Path, index: int, total: int, progress, base_url: str = MODEL_BASE_URL) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    part = dest.with_name(dest.name + ".part")
    url = base_url.rstrip("/") + "/" + spec.url_path.lstrip("/")
    progress(f"正在首次下载 {spec.label} ({index}/{total})：0%")
    hasher = hashlib.sha256()
    try:
        req = Request(url, headers={"User-Agent": "TeX2Word/1.0"})
        with urlopen(req, timeout=60) as resp, part.open("wb") as out:
            content_len = resp.headers.get("Content-Length")
            expected_bytes = int(content_len) if content_len and content_len.isdigit() else 0
            received = 0
            last_report = 0.0
            last_pct = -1
            while True:
                chunk = resp.read(1 << 20)
                if not chunk:
                    break
                out.write(chunk)
                hasher.update(chunk)
                received += len(chunk)
                now = time.monotonic()
                if expected_bytes > 0:
                    pct = min(99, int(received * 100 / expected_bytes))
                    if pct >= last_pct + 5 or (now - last_report >= 1.5 and pct != last_pct):
                        progress(
                            f"正在首次下载 {spec.label} ({index}/{total})："
                            f"{pct}%（{received / 1048576:.0f}/{expected_bytes / 1048576:.0f} MB）"
                        )
                        last_pct = pct
                        last_report = now
                elif now - last_report >= 1.5:
                    progress(f"正在首次下载 {spec.label} ({index}/{total})：{received / 1048576:.0f} MB")
                    last_report = now
        if spec.sha256 and hasher.hexdigest() != spec.sha256:
            raise ValueError(f"{spec.filename} SHA-256 校验失败，请重试。")
        part.replace(dest)
        progress(f"正在首次下载 {spec.label} ({index}/{total})：100%")
        return dest
    except BaseException as exc:
        part.unlink(missing_ok=True)
        if isinstance(exc, (KeyboardInterrupt, SystemExit)):
            raise
        raise RuntimeError(
            f"首次转换 PDF 需要下载 {spec.label}（{spec.filename}），但下载失败：{exc}。"
            f"请检查网络连接后重试，或设置环境变量 TEX2WORD_MODEL_DIR 指向本地模型目录。"
        ) from exc


def ensure_pdf_models(progress=print, *, site_packages: Path | None = None, base_url: str = MODEL_BASE_URL) -> dict[str, Path]:
    """Return local paths for the layout and formula ONNX models, downloading once if missing."""
    site = site_packages if site_packages is not None else _bundled_site_packages()
    cache = model_cache_dir()
    resolved: dict[str, Path] = {}
    missing: list[tuple[ModelSpec, Path]] = []

    for spec in PDF_MODELS:
        bundled = (site / spec.bundled_rel) if site is not None else None
        cached = cache / spec.filename
        if bundled is not None and _is_valid_cached(bundled):
            resolved[spec.key] = bundled
        elif _is_valid_cached(cached):
            resolved[spec.key] = cached
        else:
            missing.append((spec, cached))

    for index, (spec, dest) in enumerate(missing, 1):
        resolved[spec.key] = _download_model(spec, dest, index, len(missing), progress, base_url=base_url)

    return resolved
