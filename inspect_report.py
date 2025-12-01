#!/usr/bin/env python3
"""
inspect_report.py
- Reads report.csv (columns: path, full_path, file_size, has_main_resource, subresource_count, largest_subresource_bytes)
- Prints summary, top N largest files, and flags problematic filenames.
"""
import csv
import os
import sys
import unicodedata
from collections import Counter

CSV_NAME = "report.csv"
TOP_N = 20
MAX_SAFE_LEN = 200
INVALID_CHARS = set('<>:"/\\|?*')  # Windows invalid characters

def is_problematic_name(name):
    if len(name) > MAX_SAFE_LEN:
        return "TOO_LONG"
    if name.strip() != name:
        return "LEADING_TRAILING_SPACE"
    if any(c in INVALID_CHARS for c in name):
        return "INVALID_CHARS"
    if "," in name:
        return "HAS_COMMA"
    if name.count("..") > 0:
        return "DOUBLE_DOT"
    # control characters
    if any(ord(c) < 32 for c in name):
        return "CONTROL_CHAR"
    # non-normalized unicode
    if unicodedata.normalize("NFC", name) != name:
        return "NON_NFC"
    return None

def human(n):
    for unit in ('B','KB','MB','GB'):
        if n < 1024.0:
            return f"{n:.1f}{unit}"
        n /= 1024.0
    return f"{n:.1f}TB"

def main():
    if not os.path.exists(CSV_NAME):
        print(f"{CSV_NAME} not found", file=sys.stderr); sys.exit(1)
    rows = []
    with open(CSV_NAME, newline='', encoding='utf-8', errors='replace') as f:
        reader = csv.DictReader(f)
        for r in reader:
            try:
                r['file_size'] = int(r.get('file_size') or 0)
                r['subresource_count'] = int(r.get('subresource_count') or 0)
            except ValueError:
                r['file_size'] = 0
                r['subresource_count'] = 0
            rows.append(r)

    total = len(rows)
    total_bytes = sum(r['file_size'] for r in rows)
    sizes = sorted(r['file_size'] for r in rows)
    q1 = sizes[int(total*0.25)] if total>0 else 0
    q2 = sizes[int(total*0.5)] if total>0 else 0
    q3 = sizes[int(total*0.75)] if total>0 else 0

    print(f"Rows: {total}")
    print(f"Total size: {human(total_bytes)}")
    print(f"Size quartiles: Q1={human(q1)} Q2={human(q2)} Q3={human(q3)}")

    # has_main_resource distribution
    hm = Counter(r.get('has_main_resource','').lower() for r in rows)
    print("has_main_resource counts:", dict(hm))

    # subresource_count > 0
    with_subs = [r for r in rows if r['subresource_count']>0]
    print("Files reporting subresources:", len(with_subs))

    # top N largest
    top = sorted(rows, key=lambda r: r['file_size'], reverse=True)[:TOP_N]
    print(f"\nTop {TOP_N} largest files:")
    for r in top:
        print(f" {human(r['file_size']):>8}  {r['path']}")

    # problematic names
    probs = []
    for r in rows:
        basename = os.path.basename(r.get('path') or r.get('full_path') or '')
        reason = is_problematic_name(basename)
        if reason:
            probs.append((reason, r['path'], r['file_size']))
    probs_by_reason = {}
    for reason, path, size in probs:
        probs_by_reason.setdefault(reason, []).append((path, size))
    print("\nProblematic filename summary:")
    for reason, items in probs_by_reason.items():
        print(f" {reason}: {len(items)}")
        for p,s in items[:10]:
            print(f"   - {p} ({human(s)})")
    if not probs:
        print(" None detected.")

    # duplicates (same filename different path)
    names = [os.path.basename(r.get('path') or r.get('full_path') or '') for r in rows]
    dups = {n:c for n,c in Counter(names).items() if c>1}
    print("\nDuplicate basenames:", len(dups))
    for n,c in list(dups.items())[:20]:
        print(f" {n}: {c} occurrences")

if __name__ == '__main__':
    main()
