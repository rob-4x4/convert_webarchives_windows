#!/usr/bin/env python3
# convert_webarchives_windows_longpath.py
# Requirements: pip install pywebarchive
# Optional: install wkhtmltopdf for PDF output

import argparse
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Tuple

try:
    from webarchive import WebArchive
except Exception as e:
    WebArchive = None  # handled later with clearer error message

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
    "CON","PRN","AUX","NUL",
    *(f"COM{i}" for i in range(1,10)),
    *(f"LPT{i}" for i in range(1,10))
}
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
    """
    Return a normalized absolute path string.
    On Windows, optionally add the \\?\ prefix (and \\?\UNC\ for UNC paths) to support long paths.
    On non-Windows, return the absolute normalized path.
    """
    p = str(path_input)
    p = os.path.abspath(p)
    if os.name != 'nt' or not USE_LONG_PATH_PREFIX:
        return p
    # Already prefixed?
    if p.startswith('\\\\?\\'):
        return p
    # UNC: \\server\share -> \\?\UNC\server\share
    if p.startswith('\\\\'):
        return '\\\\?\\UNC\\' + p.lstrip('\\')
    return '\\\\?\\' + p

def convert_to_html(src_path: Path, html_dest: Path, dry_run: bool = False):
    ensure_parent(html_dest, dry_run=dry_run)
    if dry_run:
        return
    if WebArchive is None:
        raise RuntimeError("pywebarchive not available. Please pip install pywebarchive")
    web_src = add_long_path_prefix(str(src_path))
    web_out = add_long_path_prefix(str(html_dest))
    # WebArchive expects a path-like string; keep consistent with long-path usage
    wa = WebArchive(web_src)
    wa.to_html_file(web_out)

def html_to_pdf(html_path: Path, pdf_dest: Path, wkhtml_path: Path, dry_run: bool = False):
    ensure_parent(pdf_dest, dry_run=dry_run)
    if dry_run:
        return
    wk = str(wkhtml_path) if wkhtml_path else None
    if not wk or not os.path.isfile(wk):
        raise FileNotFoundError("wkhtmltopdf not configured or not found at: " + str(wkhtml_path))
    html_arg = add_long_path_prefix(str(html_path))
    pdf_arg = add_long_path_prefix(str(pdf_dest))
    # Use subprocess with check=True to raise on errors
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

def process_single_file(src_file: Path, out_html_root: Path, out_pdf_root: Path,
                        failed_root: Path, wkhtml_path: Path,
                        skip_pdf: bool, dry_run: bool) -> Tuple[bool, str]:
    """
    Convert a single .webarchive file. Returns (success, message).
    """
    try:
        if not src_file.exists():
            return False, "source file not found"
        rel_root = Path(".")
        safe_stem, _ = safe_stem_with_ext(src_file.name)
        out_html = out_html_root / rel_root / (safe_stem + ".html")
        out_pdf = out_pdf_root / rel_root / (safe_stem + ".pdf")
        print(f"Converting: {src_file}")
        convert_to_html(src_file, out_html, dry_run=dry_run)
        print(f"    HTML -> {out_html}")
        if (not skip_pdf) and wkhtml_path:
            try:
                html_to_pdf(out_html, out_pdf, wkhtml_path, dry_run=dry_run)
                print(f"    PDF  -> {out_pdf}")
            except Exception as e:
                return False, f"PDF conversion failed: {e}"
        return True, "OK"
    except Exception as e:
        # Attempt to move failed file
        try:
            move_failed(src_file, failed_root, rel_root, dry_run=dry_run)
            print(f"    Moved failed .webarchive to {failed_root / rel_root}")
        except Exception as me:
            return False, f"Conversion error: {e}; Additionally failed to move: {me}"
        return False, f"Conversion error: {e}"

def walk_and_process(src_root: Path, out_html_root: Path, out_pdf_root: Path,
                     failed_root: Path, wkhtml_path: Path,
                     skip_pdf: bool, dry_run: bool):
    if not src_root.exists():
        print("Source folder not found:", src_root)
        return
    out_html_root.mkdir(parents=True, exist_ok=True)
    out_pdf_root.mkdir(parents=True, exist_ok=True)
    failed_root.mkdir(parents=True, exist_ok=True)

    count = 0
    errors = []

    for root, dirs, files in os.walk(src_root):
        root_path = Path(root)
        try:
            rel_root = root_path.relative_to(src_root)
        except Exception:
            rel_root = Path(".")
        rel_root_trunc = truncate_rel_parts(rel_root)

        for fname in files:
            if not fname.lower().endswith('.webarchive'):
                continue
            count += 1
            src_file = root_path / fname
            safe_stem, _ = safe_stem_with_ext(src_file.name)
            out_html = out_html_root / rel_root_trunc / (safe_stem + ".html")
            out_pdf = out_pdf_root / rel_root_trunc / (safe_stem + ".pdf")

            try:
                print(f"[{count}] Converting: {src_file}")
                convert_to_html(src_file, out_html, dry_run=dry_run)
                print(f"    HTML -> {out_html}")
                if (not skip_pdf) and wkhtml_path:
                    try:
                        html_to_pdf(out_html, out_pdf, wkhtml_path, dry_run=dry_run)
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
    print(f"Total .webarchive files found: {count}")
    if errors:
        print(f"Total items with errors: {len(errors)}")
        for fpath, err in errors:
            print(f" - {fpath}: {err}")
    else:
        print("No errors encountered.")

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
    return p.parse_args()

def main():
    args = parse_args()

    if args.test_file:
        # Single-file test mode: convert the provided file only
        print("Running single-file test for:", args.test_file)
        success, message = process_single_file(
            src_file=args.test_file,
            out_html_root=args.out_html,
            out_pdf_root=args.out_pdf,
            failed_root=args.failed,
            wkhtml_path=args.wkhtml,
            skip_pdf=args.skip_pdf,
            dry_run=args.dry_run
        )
        if success:
            print("Test conversion succeeded:", message)
            return 0
        else:
            print("Test conversion failed:", message)
            return 2

    # Normal batch mode: walk the tree
    try:
        walk_and_process(
            src_root=args.src,
            out_html_root=args.out_html,
            out_pdf_root=args.out_pdf,
            failed_root=args.failed,
            wkhtml_path=args.wkhtml,
            skip_pdf=args.skip_pdf,
            dry_run=args.dry_run
        )
        return 0
    except Exception as e:
        print("Fatal error:", e)
        return 1

if __name__ == "__main__":
    # Informative note when running on non-Windows
    if os.name != 'nt':
        print("Note: This script is written for Windows and uses Windows long-path prefixes when enabled.")
    sys.exit(main())
