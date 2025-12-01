#!/usr/bin/env python3
# scan_webarchives_report.py
# Usage: pip install pywebarchive
# Example:
#   python scan_webarchives_report.py --src "D:\usbwork" --out report.csv

import csv
import os
import sys
import tempfile
from pathlib import Path
from typing import Tuple, Optional

try:
    from webarchive import WebArchive
except Exception:
    WebArchive = None

# Windows long-path helper (keeps behavior consistent with your other scripts)
USE_LONG_PATH_PREFIX = True
def add_long_path_prefix(p: str) -> str:
    p = os.path.abspath(p)
    if os.name != 'nt' or not USE_LONG_PATH_PREFIX:
        return p
    if p.startswith('\\\\?\\'):
        return p
    if p.startswith('\\\\'):
        return '\\\\?\\UNC\\' + p.lstrip('\\')
    return '\\\\?\\' + p

def inspect_webarchive(path: Path) -> Tuple[bool, int, Optional[int]]:
    """
    Returns (has_main_resource, subresource_count, largest_subresource_bytes or None)
    If pywebarchive is not available, raises RuntimeError.
    """
    if WebArchive is None:
        raise RuntimeError("pywebarchive not installed. Install with: pip install pywebarchive")

    try:
        wa = WebArchive(add_long_path_prefix(str(path)))
    except Exception as e:
        # treat as unreadable
        return False, 0, None

    def maybe_attr(obj, *names):
        for n in names:
            if hasattr(obj, n):
                try:
                    return getattr(obj, n)
                except Exception:
                    pass
        return None

    mr = maybe_attr(wa, "main_resource", "_main_resource", "web_main_resource")
    has_main = mr is not None

    # gather subresources from several likely attributes
    sub_count = 0
    largest = 0
    found_any = False
    for attr in ("subresources", "_subresources", "resources", "web_resources", "WebResources", "resource_count", "subframe_archives"):
        if hasattr(wa, attr):
            try:
                col = getattr(wa, attr)
            except Exception:
                continue
            if col is None:
                continue
            # attempt to iterate safely
            try:
                items = list(col) if not isinstance(col, int) else []
            except Exception:
                # fallback: try calling get_subresource / resource_count if available
                items = []
            if items:
                for r in items:
                    found_any = True
                    sub_count += 1
                    # probe for bytes-like fields
                    size = None
                    for key in ("data", "data_bytes", "content", "html", "value"):
                        if hasattr(r, key):
                            try:
                                v = getattr(r, key)
                                if isinstance(v, (bytes, bytearray)):
                                    size = len(v); break
                                if isinstance(v, str):
                                    size = len(v); break
                            except Exception:
                                pass
                    if size is not None and size > largest:
                        largest = size
    if not found_any:
        # if library exposes resource_count and get_subresource, try those
        try:
            rc = getattr(wa, "resource_count", None)
            if isinstance(rc, int) and rc > 0:
                sub_count = rc
                # try to sample several subresources to find sizes
                for i in range(min(10, rc)):
                    try:
                        r = wa.get_subresource(i)
                        size = None
                        for key in ("data", "data_bytes", "content", "html", "value"):
                            if hasattr(r, key):
                                v = getattr(r, key)
                                if isinstance(v, (bytes, bytearray)):
                                    size = len(v); break
                                if isinstance(v, str):
                                    size = len(v); break
                        if size is not None and size > largest:
                            largest = size
                    except Exception:
                        pass
        except Exception:
            pass

    if largest == 0:
        largest = None

    return has_main, sub_count, largest

def scan_folder(src: Path, out_csv: Path):
    src = src.resolve()
    rows = []
    for root, dirs, files in os.walk(src):
        for fn in files:
            if fn.lower().endswith(".webarchive"):
                p = Path(root) / fn
                try:
                    size = p.stat().st_size
                except Exception:
                    size = None
                try:
                    has_main, sub_count, largest = inspect_webarchive(p)
                except Exception as e:
                    has_main, sub_count, largest = False, 0, None
                rows.append({
                    "path": str(p.relative_to(src)),
                    "full_path": str(p),
                    "file_size": size if size is not None else "",
                    "has_main_resource": "yes" if has_main else "no",
                    "subresource_count": sub_count,
                    "largest_subresource_bytes": largest if largest is not None else ""
                })
    # write CSV
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with open(out_csv, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["path","full_path","file_size","has_main_resource","subresource_count","largest_subresource_bytes"])
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    print(f"Wrote {len(rows)} records to {out_csv}")

def parse_args():
    import argparse
    p = argparse.ArgumentParser(prog="scan_webarchives_report.py", description="Scan folder of .webarchive files and emit CSV report")
    p.add_argument("--src", type=Path, required=True, help="Root folder to scan")
    p.add_argument("--out", type=Path, default=Path.cwd() / "webarchive_report.csv", help="Output CSV path")
    return p.parse_args()

def main():
    args = parse_args()
    if WebArchive is None:
        print("pywebarchive not installed. Install with: pip install pywebarchive", file=sys.stderr)
        sys.exit(2)
    scan_folder(args.src, args.out)

if __name__ == "__main__":
    main()
