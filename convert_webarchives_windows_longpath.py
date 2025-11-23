#!/usr/bin/env python3
# convert_webarchives_windows_longpath.py
# Requirements: pip install pywebarchive
# Optional: wkhtmltopdf for PDF output

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from webarchive import WebArchive

# ========== CONFIG ==========
SRC = Path(r"D:\usbwork")
OUT_HTML = Path(r"D:\USBhtml")
OUT_PDF = Path(r"D:\USBpdf")
FAILED = Path(r"D:\USBfailed")
# Path to wkhtmltopdf executable or set to None to skip PDF conversion
WKHTMLTOPDF = Path(r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
# Maximum allowed length for individual path components after sanitization
MAX_COMPONENT_LEN = 100
# Maximum allowed length for filename stem (before extension)
MAX_STEM_LEN = 200
# If True, prefix absolute paths with \\?\ to enable long path handling on Windows
USE_LONG_PATH_PREFIX = True
# ============================

INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
RESERVED_NAMES = {
    "CON","PRN","AUX","NUL",
    *(f"COM{i}" for i in range(1,10)),
    *(f"LPT{i}" for i in range(1,10))
}

def safe_component(name: str, max_len: int = MAX_COMPONENT_LEN) -> str:
    """
    Make a path component safe for Windows:
    - Replace invalid characters with underscore
    - Trim trailing spaces/dots
    - Avoid reserved names
    - Truncate to max_len
    """
    # Keep extension handling to callers when needed
    name = INVALID_CHARS_RE.sub("_", name)
    name = name.rstrip(" .")
    if not name:
        name = "_"
    if name.upper() in RESERVED_NAMES:
        name = "_" + name
    if len(name) > max_len:
        name = name[:max_len]
    return name

def safe_stem_with_ext(orig_name: str, stem_max: int = MAX_STEM_LEN) -> (str, str):
    """
    Return (safe_stem, extension) for an original filename.
    """
    stem, ext = os.path.splitext(orig_name)
    stem = INVALID_CHARS_RE.sub("_", stem).rstrip(" .")
    if not stem:
        stem = "_"
    if stem.upper() in RESERVED_NAMES:
        stem = "_" + stem
    if len(stem) > stem_max:
        stem = stem[:stem_max]
    # Clean extension but keep leading dot
    ext = ext if ext else ""
    return stem, ext

def ensure_parent(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)

def add_long_path_prefix(path_str: str) -> str:
    """
    Add Windows long-path prefix \\?\ to absolute paths when USE_LONG_PATH_PREFIX is True.
    This function expects an absolute Windows path.
    """
    if not USE_LONG_PATH_PREFIX:
        return path_str
    # Normalize slashes and absolute path
    p = os.path.abspath(path_str)
    # Already prefixed?
    if p.startswith('\\\\?\\'):
        return p
    # UNC paths need special handling: \\server\share -> \\?\UNC\server\share
    if p.startswith('\\\\'):
        return '\\\\?\\UNC\\' + p.lstrip('\\')
    return '\\\\?\\' + p

def convert_to_html(src_path: Path, html_dest: Path):
    ensure_parent(html_dest)
    # webarchive accepts a normal path string; pass long-prefixed path for long path support
    web_src = add_long_path_prefix(str(src_path))
    web_out = add_long_path_prefix(str(html_dest))
    wa = WebArchive(web_src)
    wa.to_html_file(web_out)

def html_to_pdf(html_path: Path, pdf_dest: Path):
    ensure_parent(pdf_dest)
    if not WKHTMLTOPDF or not WKHTMLTOPDF.exists():
        raise FileNotFoundError("wkhtmltopdf not configured or not found")
    # Use long-path-prefixed strings for subprocess arguments on Windows
    html_arg = add_long_path_prefix(str(html_path))
    pdf_arg = add_long_path_prefix(str(pdf_dest))
    subprocess.run([str(WKHTMLTOPDF), '--quiet', html_arg, pdf_arg], check=True)

def move_failed(src_file: Path, rel_root: Path):
    dest = FAILED / rel_root / src_file.name
    dest.parent.mkdir(parents=True, exist_ok=True)
    try:
        # Use shutil.move which works with long prefix when paths are passed as strings
        shutil.move(add_long_path_prefix(str(src_file)), add_long_path_prefix(str(dest)))
    except Exception:
        # Fallback: copy then remove
        try:
            shutil.copy2(add_long_path_prefix(str(src_file)), add_long_path_prefix(str(dest)))
            os.remove(add_long_path_prefix(str(src_file)))
        except Exception as e2:
            print(f"    Failed to move or copy failed file {src_file}: {e2}")

def truncate_rel_parts(rel: Path) -> Path:
    """
    Truncate each component in a relative path to safe_component limits to avoid extremely long paths.
    """
    parts = []
    for part in rel.parts:
        parts.append(safe_component(part, MAX_COMPONENT_LEN))
    return Path(*parts) if parts else Path("")

def main():
    if not SRC.exists():
        print("Source folder not found:", SRC)
        return

    OUT_HTML.mkdir(parents=True, exist_ok=True)
    OUT_PDF.mkdir(parents=True, exist_ok=True)
    FAILED.mkdir(parents=True, exist_ok=True)

    count = 0
    errors = []

    for root, dirs, files in os.walk(SRC):
        root_path = Path(root)
        try:
            rel_root = root_path.relative_to(SRC)
        except Exception:
            rel_root = Path(".")
        rel_root_trunc = truncate_rel_parts(rel_root)

        for fname in files:
            if not fname.lower().endswith('.webarchive'):
                continue
            count += 1
            src_file = root_path / fname

            safe_stem, _ = safe_stem_with_ext(src_file.name)
            # build output filenames (use .html and .pdf extensions)
            out_html = OUT_HTML / rel_root_trunc / (safe_stem + ".html")
            out_pdf = OUT_PDF / rel_root_trunc / (safe_stem + ".pdf")

            try:
                print(f"[{count}] Converting: {src_file}")
                convert_to_html(src_file, out_html)
                print(f"    HTML -> {out_html}")

                if WKHTMLTOPDF and WKHTMLTOPDF.exists():
                    try:
                        html_to_pdf(out_html, out_pdf)
                        print(f"    PDF  -> {out_pdf}")
                    except Exception as e:
                        print(f"    PDF conversion failed for {out_html}: {e}")
                        errors.append((str(src_file), f"PDF error: {e}"))
                # keep original .webarchive in place on success
            except Exception as e:
                print(f"    ERROR converting {src_file}: {e}")
                errors.append((str(src_file), str(e)))
                try:
                    move_failed(src_file, rel_root_trunc)
                    print(f"    Moved failed .webarchive to {FAILED / rel_root_trunc}")
                except Exception as me:
                    print(f"    Failed to move failed file {src_file}: {me}")

    print(f"\nDone. Scanned tree starting at: {SRC}")
    print(f"Total .webarchive files found: {count}")
    if errors:
        print(f"Total items with errors: {len(errors)}")
        for fpath, err in errors:
            print(f" - {fpath}: {err}")
    else:
        print("No errors encountered.")

if __name__ == "__main__":
    # Ensure script is running on Windows for long-path expectations, but it will run elsewhere too.
    if os.name != 'nt':
        print("Note: This script is written for Windows (nt). Some long-path behaviors are Windows-specific.")
    main()
