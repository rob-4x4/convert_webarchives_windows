#!/usr/bin/env python3
# convert_webarchives_windows_longpath.py
# Requirements: pip install pywebarchive tqdm
# Optional: install wkhtmltopdf for PDF output

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Tuple, List

try:
    from webarchive import WebArchive
except Exception:
    WebArchive = None

try:
    from tqdm import tqdm
except Exception:
    tqdm = None

# ========== DEFAULT CONFIG ==========
DEFAULT_SRC = Path(r"D:\usbwork")
DEFAULT_OUT_HTML = Path(r"D:\USBhtml")
DEFAULT_OUT_PDF = Path(r"D:\USBpdf")
DEFAULT_FAILED = Path(r"D:\USBfailed")
DEFAULT_WKHTMLTOPDF = Path(r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
MAX_COMPONENT_LEN = 100
MAX_STEM_LEN = 200
USE_LONG_PATH_PREFIX = True
INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10))
}
MIN_WEBARCHIVE_SIZE = 128
# ============================


def safe_component(name: str, max_len: int = MAX_COMPONENT_LEN) -> str:
    name = INVALID_CHARS_RE.sub("_", name)
    name = name.rstrip(" .")
    if not name:
        name = "_"
    if name.upper() in RESERVED_NAMES:
        name = "_" + name
    if len(name) > max_len:
        name = name[:max_len]
    return name


def safe_stem_with_ext(orig_name: str, stem_max: int = MAX_STEM_LEN) -> Tuple[str, str]:
    stem, ext = os.path.splitext(orig_name)
    stem = INVALID_CHARS_RE.sub("_", stem).rstrip(" .")
    if not stem:
        stem = "_"
    if stem.upper() in RESERVED_NAMES:
        stem = "_" + stem
    if len(stem) > stem_max:
        stem = stem[:stem_max]
    ext = ext if ext else ""
    return stem, ext


def ensure_parent(path: Path, dry_run: bool = False):
    if dry_run:
        return
    path.parent.mkdir(parents=True, exist_ok=True)


def add_long_path_prefix(path_input: str) -> str:
    r"""
    Return a normalized absolute path string.
    On Windows, optionally add the \\?\ prefix (and \\?\UNC\ for UNC paths) to support long paths.
    """
    p = str(path_input)
    p = os.path.abspath(p)
    if os.name != 'nt' or not USE_LONG_PATH_PREFIX:
        return p
    if p.startswith('\\\\?\\'):
        return p
    if p.startswith('\\\\'):
        return '\\\\?\\UNC\\' + p.lstrip('\\')
    return '\\\\?\\' + p


def _write_text_file(path: str, text: str):
    # write text to file path (path may be long-path-prefixed string)
    with open(path, "w", encoding="utf-8", errors="replace") as fh:
        fh.write(text)


# ----------------- Helpers to build index / extract resources ----------------- #
def _build_index_from_wa(wa, src_name: str) -> str:
    parts = [
        "<!doctype html>",
        "<html><head><meta charset='utf-8'>",
        f"<title>Index of {src_name}</title>",
        "</head><body>",
        f"<h1>Index of resources in {src_name}</h1>",
        "<ul>"
    ]

    try:
        mr = getattr(wa, "main_resource", None) or getattr(wa, "_main_resource", None)
        if mr:
            name = getattr(mr, "url", None) or getattr(mr, "URL", None) or getattr(mr, "filename", None) or "main"
            mime = getattr(mr, "mimeType", None) or getattr(mr, "MIMEType", None) or None
            parts.append(f"<li>Main resource: {name} (mime={mime})</li>")
    except Exception:
        pass

    try:
        subs = None
        for attr in ("subresources", "_subresources", "subframe_archives", "subframe"):
            if hasattr(wa, attr):
                try:
                    col = getattr(wa, attr)
                    if col:
                        subs = list(col) if isinstance(col, (list, tuple)) else list(col)
                        break
                except Exception:
                    continue
        if subs:
            for i, r in enumerate(subs, start=1):
                rname = getattr(r, "url", None) or getattr(r, "URL", None) or getattr(r, "filename", None) or f"resource_{i}"
                rmime = getattr(r, "mimeType", None) or getattr(r, "MIMEType", None) or None
                parts.append(f"<li>Resource: {rname} (mime={rmime})</li>")
    except Exception:
        pass

    parts.append("</ul></body></html>")
    return "\n".join(parts)


def _guess_html_resource_from_wa(wa) -> Tuple[str, str]:
    candidates = []
    for attr in ("web_main_resource", "main_resource", "mainResource", "WebMainResource", "_main_resource"):
        if hasattr(wa, attr):
            mr = getattr(wa, attr)
            if mr:
                candidates.append(mr)
    for attr in ("resources", "web_resources", "WebResources", "resourcesList", "subresources", "_subresources"):
        if hasattr(wa, attr):
            rcol = getattr(wa, attr)
            if rcol:
                try:
                    if isinstance(rcol, (list, tuple)):
                        for r in rcol:
                            candidates.append(r)
                    else:
                        for r in rcol:
                            candidates.append(r)
                except Exception:
                    pass

    def inspect_resource_obj(obj):
        mime = None
        data = None
        for mname in ("mimeType", "MIMEType", "mime", "contentType"):
            if hasattr(obj, mname):
                try:
                    mime = getattr(obj, mname)
                    break
                except Exception:
                    pass
        for dname in ("data", "data_bytes", "content", "html", "value"):
            if hasattr(obj, dname):
                try:
                    data = getattr(obj, dname)
                    break
                except Exception:
                    pass
        if mime is None and isinstance(obj, dict):
            for key in ("mimeType", "MIMEType", "mime", "contentType"):
                if key in obj:
                    mime = obj.get(key)
                    break
        if data is None and isinstance(obj, dict):
            for key in ("data", "content", "html", "value"):
                if key in obj:
                    data = obj.get(key)
                    break
        if data is not None and isinstance(data, (bytes, bytearray)):
            try:
                text = data.decode("utf-8")
            except Exception:
                try:
                    text = data.decode("latin-1")
                except Exception:
                    text = None
            data = text
        return mime, data

    for res in candidates:
        try:
            mime, text = inspect_resource_obj(res)
            if mime and isinstance(mime, str) and "html" in mime.lower() and text:
                return mime, text
            if isinstance(res, str) and "<html" in res.lower():
                return "text/html", res
        except Exception:
            continue
    return None, None


# ----------------- Conversion using pywebarchive 0.5.2 APIs ----------------- #
def convert_to_html(src_path: Path, html_dest: Path, dry_run: bool = False):
    """
    Uses WebArchive.to_html() when available; falls back to extracting main_resource
    or building an index if no main resource is present.
    """
    ensure_parent(html_dest, dry_run=dry_run)
    if dry_run:
        return
    if WebArchive is None:
        raise RuntimeError("pywebarchive not installed; pip install pywebarchive")

    web_src = add_long_path_prefix(str(src_path))
    web_out = add_long_path_prefix(str(html_dest))

    wa = WebArchive(web_src)

    last_exc = None
    if hasattr(wa, "to_html"):
        try:
            html_text = wa.to_html()
            if html_text:
                _write_text_file(web_out, html_text)
                return
        except Exception as e:
            last_exc = e

    try:
        mime, html_text = _guess_html_resource_from_wa(wa)
        if mime and html_text:
            _write_text_file(web_out, html_text)
            return
    except Exception as e:
        last_exc = e

    try:
        mr = getattr(wa, "main_resource", None) or getattr(wa, "_main_resource", None)
        if mr:
            data = None
            for key in ("data", "data_bytes", "content", "html"):
                if hasattr(mr, key):
                    try:
                        data = getattr(mr, key)
                        break
                    except Exception:
                        continue
            if isinstance(data, (bytes, bytearray)):
                try:
                    text = data.decode("utf-8")
                except Exception:
                    try:
                        text = data.decode("latin-1")
                    except Exception:
                        text = None
                if text:
                    _write_text_file(web_out, text)
                    return
            if isinstance(data, str):
                _write_text_file(web_out, data)
                return
    except Exception:
        pass

    idx = _build_index_from_wa(wa, src_path.name)
    _write_text_file(web_out, idx)
    return


def html_to_pdf(html_path: Path, pdf_dest: Path, wkhtml_path: Path, dry_run: bool = False):
    ensure_parent(pdf_dest, dry_run=dry_run)
    if dry_run:
        return
    wk = str(wkhtml_path) if wkhtml_path else None
    if not wk or not os.path.isfile(wk):
        raise FileNotFoundError("wkhtmltopdf not configured or not found at: " + str(wkhtml_path))
    html_arg = add_long_path_prefix(str(html_path))
    pdf_arg = add_long_path_prefix(str(pdf_dest))
    subprocess.run([wk, '--quiet', html_arg, pdf_arg], check=True)


def move_failed(src_file: Path, failed_root: Path, rel_root: Path, dry_run: bool = False):
    dest = failed_root / rel_root / src_file.name
    if dry_run:
        print(f"    [dry-run] would move {src_file} -> {dest}")
        return
    dest.parent.mkdir(parents=True, exist_ok=True)
    s = add_long_path_prefix(str(src_file))
    d = add_long_path_prefix(str(dest))
    try:
        shutil.move(s, d)
    except Exception:
        try:
            shutil.copy2(s, d)
            os.remove(s)
        except Exception as e2:
            raise RuntimeError(f"Failed to move or copy failed file {src_file}: {e2}") from e2


def truncate_rel_parts(rel: Path) -> Path:
    parts = []
    for part in rel.parts:
        parts.append(safe_component(part, MAX_COMPONENT_LEN))
    return Path(*parts) if parts else Path("")


# ---------- Validation helpers ----------

def is_likely_webarchive_file(path: Path, min_size: int = MIN_WEBARCHIVE_SIZE) -> bool:
    if path.name.startswith("._"):
        return False
    if not path.is_file():
        return False
    if path.suffix.lower() != ".webarchive":
        return False
    try:
        size = path.stat().st_size
    except Exception:
        return False
    if size < min_size:
        return False
    try:
        with path.open("rb") as fh:
            chunk = fh.read(8192)
    except Exception:
        return False
    markers = [b'<?xml', b'plist', b'AppleWebArchive', b'WebResource', b'WebMainResource', b'bplist00']
    return any(m in chunk for m in markers)


def is_valid_webarchive_by_parsing(path: Path, dry_run: bool = False) -> Tuple[bool, str]:
    """
    Definitive parsing validation using pywebarchive.
    Creates a temp file with mkstemp then closes descriptor so Windows won't block it.
    """
    if WebArchive is None:
        return False, "pywebarchive not installed"
    if dry_run:
        return True, "dry-run: parsing skipped"

    try:
        wa = WebArchive(add_long_path_prefix(str(path)))
    except Exception as e:
        return False, f"pywebarchive init failed: {e}"

    fd = None
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(prefix="wa_check_", suffix=".html")
        os.close(fd)
        tmp_out = add_long_path_prefix(tmp_path)

        try:
            if hasattr(wa, "to_html"):
                try:
                    text = wa.to_html()
                    if text:
                        _write_text_file(tmp_out, text)
                        return True, ""
                except Exception:
                    pass
        except Exception:
            pass

        try:
            mr = getattr(wa, "main_resource", None) or getattr(wa, "_main_resource", None)
            if mr:
                data = None
                for key in ("data", "data_bytes", "content", "html"):
                    if hasattr(mr, key):
                        try:
                            data = getattr(mr, key)
                            break
                        except Exception:
                            continue
                if isinstance(data, (bytes, bytearray)):
                    try:
                        text = data.decode("utf-8")
                    except Exception:
                        try:
                            text = data.decode("latin-1")
                        except Exception:
                            text = None
                    if text:
                        _write_text_file(tmp_out, text)
                        return True, ""
                if isinstance(data, str):
                    _write_text_file(tmp_out, data)
                    return True, ""
        except Exception:
            pass

        try:
            idx = _build_index_from_wa(wa, path.name)
            _write_text_file(tmp_out, idx)
            return True, ""
        except Exception as e:
            return False, f"fallback index generation failed: {e}"

    except Exception as e:
        return False, f"temp file error: {e}"
    finally:
        try:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


# ---------- File discovery and processing ----------

def gather_webarchive_files(src_root: Path) -> List[Path]:
    files: List[Path] = []
    for root, dirs, filenames in os.walk(src_root):
        for fn in filenames:
            if fn.lower().endswith('.webarchive'):
                files.append(Path(root) / fn)
    return files


def clean_sidecars_to_failed(src_root: Path, failed_root: Path, dry_run: bool = False) -> int:
    moved = 0
    for root, dirs, filenames in os.walk(src_root):
        for fn in filenames:
            if fn.startswith("._") and fn.lower().endswith(".webarchive"):
                src_file = Path(root) / fn
                try:
                    try:
                        rel = Path(root).relative_to(src_root)
                    except Exception:
                        rel = Path(".")
                    move_failed(src_file, failed_root, rel, dry_run=dry_run)
                    moved += 1
                except Exception as e:
                    print(f"    Failed to move sidecar {src_file}: {e}")
    return moved


def process_single_file(src_file: Path, out_html_root: Path, out_pdf_root: Path,
                        failed_root: Path, wkhtml_path: Path,
                        skip_pdf: bool, dry_run: bool, validate: bool) -> Tuple[bool, str]:
    try:
        if not src_file.exists():
            return False, "source file not found"
        if not is_likely_webarchive_file(src_file):
            return False, "heuristic validation failed"
        if validate:
            valid, msg = is_valid_webarchive_by_parsing(src_file, dry_run=dry_run)
            if not valid:
                return False, f"parsing validation failed: {msg}"
        rel_root = Path(".")
        safe_stem, _ = safe_stem_with_ext(src_file.name)
        out_html = out_html_root / rel_root / (safe_stem + ".html")
        out_pdf = out_pdf_root / rel_root / (safe_stem + ".pdf")
        convert_to_html(src_file, out_html, dry_run=dry_run)
        if (not skip_pdf) and wkhtml_path:
            try:
                html_to_pdf(out_html, out_pdf, wkhtml_path, dry_run=dry_run)
            except Exception as e:
                return False, f"PDF conversion failed: {e}"
        return True, "OK"
    except Exception as e:
        try:
            move_failed(src_file, failed_root, Path("."), dry_run=dry_run)
            return False, f"Conversion error: {e}; moved to failed"
        except Exception as me:
            return False, f"Conversion error: {e}; additionally failed to move: {me}"


def walk_and_process(src_root: Path, out_html_root: Path, out_pdf_root: Path,
                     failed_root: Path, wkhtml_path: Path,
                     skip_pdf: bool, dry_run: bool, use_progress: bool, validate: bool,
                     clean_sidecars: bool):
    if not src_root.exists():
        print("Source folder not found:", src_root)
        return

    out_html_root.mkdir(parents=True, exist_ok=True)
    out_pdf_root.mkdir(parents=True, exist_ok=True)
    failed_root.mkdir(parents=True, exist_ok=True)

    if clean_sidecars:
        print("Cleaning AppleDouble sidecar files (._*.webarchive) into failed folder...")
        moved = clean_sidecars_to_failed(src_root, failed_root, dry_run=dry_run)
        print(f"  Sidecars moved (or planned): {moved}")

    all_files = gather_webarchive_files(src_root)
    total = len(all_files)
    if total == 0:
        print("No .webarchive files found under:", src_root)
        return

    print(f"Found {total} .webarchive files. Starting conversion{' (dry-run)' if dry_run else ''}.")

    iterator = all_files
    if use_progress:
        if tqdm is None:
            print("tqdm not installed; install it with: pip install tqdm")
            iterator = all_files
        else:
            iterator = tqdm(all_files, desc="Converting", unit="file", ncols=100)

    errors = []
    count = 0

    for src_file in iterator:
        count += 1
        try:
            rel_root = src_file.parent.relative_to(src_root)
        except Exception:
            rel_root = Path(".")
        rel_root_trunc = truncate_rel_parts(rel_root)

        safe_stem, _ = safe_stem_with_ext(src_file.name)
        out_html = out_html_root / rel_root_trunc / (safe_stem + ".html")
        out_pdf = out_pdf_root / rel_root_trunc / (safe_stem + ".pdf")

        try:
            if not dry_run:
                print(f"[{count}/{total}] Converting: {src_file}")
            else:
                print(f"[{count}/{total}] (dry-run) Would convert: {src_file}")
            if not is_likely_webarchive_file(src_file):
                print(f"    Skipping (heuristic): {src_file}")
                errors.append((str(src_file), "heuristic validation failed"))
                try:
                    move_failed(src_file, failed_root, rel_root_trunc, dry_run=dry_run)
                    print(f"    Moved heuristic-failed .webarchive to {failed_root / rel_root_trunc}")
                except Exception as me:
                    print(f"    Failed to move heuristic-failed file {src_file}: {me}")
                continue
            if validate:
                valid, msg = is_valid_webarchive_by_parsing(src_file, dry_run=dry_run)
                if not valid:
                    print(f"    Parsing validation failed for {src_file}: {msg}")
                    errors.append((str(src_file), f"parsing validation failed: {msg}"))
                    try:
                        move_failed(src_file, failed_root, rel_root_trunc, dry_run=dry_run)
                        print(f"    Moved invalid .webarchive to {failed_root / rel_root_trunc}")
                    except Exception as me:
                        print(f"    Failed to move invalid file {src_file}: {me}")
                    continue

            convert_to_html(src_file, out_html, dry_run=dry_run)
            if not dry_run:
                print(f"    HTML -> {out_html}")
            if (not skip_pdf) and wkhtml_path:
                try:
                    html_to_pdf(out_html, out_pdf, wkhtml_path, dry_run=dry_run)
                    if not dry_run:
                        print(f"    PDF  -> {out_pdf}")
                except Exception as e:
                    print(f"    PDF conversion failed for {out_html}: {e}")
                    errors.append((str(src_file), f"PDF error: {e}"))
        except Exception as e:
            print(f"    ERROR converting {src_file}: {e}")
            errors.append((str(src_file), str(e)))
            try:
                move_failed(src_file, failed_root, rel_root_trunc, dry_run=dry_run)
                print(f"    Moved failed .webarchive to {failed_root / rel_root_trunc}")
            except Exception as me:
                print(f"    Failed to move failed file {src_file}: {me}")

    print(f"\nDone. Scanned tree starting at: {src_root}")
    print(f"Total .webarchive files processed: {total}")
    if errors:
        print(f"Total items with errors: {len(errors)}")
        for fpath, err in errors:
            print(f" - {fpath}: {err}")
    else:
        print("No errors encountered.")


# ---------- CLI ----------

def parse_args():
    p = argparse.ArgumentParser(description="Convert .webarchive files to HTML (and PDF via wkhtmltopdf)")
    p.add_argument("--src", type=Path, default=DEFAULT_SRC, help="Source folder to scan")
    p.add_argument("--out-html", type=Path, default=DEFAULT_OUT_HTML, help="Output root for HTML files")
    p.add_argument("--out-pdf", type=Path, default=DEFAULT_OUT_PDF, help="Output root for PDF files")
    p.add_argument("--failed", type=Path, default=DEFAULT_FAILED, help="Folder to move failed .webarchive files")
    p.add_argument("--wkhtml", type=Path, default=DEFAULT_WKHTMLTOPDF, help="Path to wkhtmltopdf executable")
    p.add_argument("--skip-pdf", action="store_true", help="Skip PDF conversion")
    p.add_argument("--dry-run", action="store_true", help="Do not write files; print what would be done")
    p.add_argument("--test-file", type=Path, help="Test single .webarchive file and exit")
    p.add_argument("--no-progress", action="store_true", help="Disable progress bar")
    p.add_argument("--validate", action="store_true", help="Enable parsing validation with pywebarchive (slower)")
    p.add_argument("--clean-sidecars", action="store_true", help="Move AppleDouble sidecar files (._*.webarchive) to FAILED before processing")
    return p.parse_args()


def main():
    args = parse_args()
    use_progress = not args.no_progress

    if args.test_file:
        print("Running single-file test for:", args.test_file)
        success, message = process_single_file(
            src_file=args.test_file,
            out_html_root=args.out_html,
            out_pdf_root=args.out_pdf,
            failed_root=args.failed,
            wkhtml_path=args.wkhtml,
            skip_pdf=args.skip_pdf,
            dry_run=args.dry_run,
            validate=args.validate
        )
        if success:
            print("Test conversion succeeded:", message)
            return 0
        else:
            print("Test conversion failed:", message)
            return 2

    try:
        walk_and_process(
            src_root=args.src,
            out_html_root=args.out_html,
            out_pdf_root=args.out_pdf,
            failed_root=args.failed,
            wkhtml_path=args.wkhtml,
            skip_pdf=args.skip_pdf,
            dry_run=args.dry_run,
            use_progress=use_progress,
            validate=args.validate,
            clean_sidecars=args.clean_sidecars
        )
        return 0
    except Exception as e:
        print("Fatal error:", e)
        return 1


if __name__ == "__main__":
    if os.name != 'nt':
        print("Note: This script is written for Windows and uses Windows long-path prefixes when enabled.")
    sys.exit(main())
