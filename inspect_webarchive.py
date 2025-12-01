# inspect_webarchive.py
# pip install pywebarchive
from pathlib import Path
from webarchive import WebArchive

def probe(wa):
    def maybe_attr(obj, *names):
        for n in names:
            if hasattr(obj, n):
                try:
                    return getattr(obj, n)
                except Exception:
                    pass
        return None

    print("Top-level attributes:", [a for a in dir(wa) if not a.startswith("_")][:50])
    mr = maybe_attr(wa, "main_resource", "_main_resource", "web_main_resource")
    if mr:
        print("Main resource:")
        print("  url:", maybe_attr(mr, "url", "URL", "filename", "path"))
        print("  mime:", maybe_attr(mr, "mimeType", "MIMEType", "contentType"))
        # try to find data bytes
        for key in ("data", "data_bytes", "content", "html", "value"):
            if hasattr(mr, key):
                val = getattr(mr, key)
                if isinstance(val, (bytes, bytearray)):
                    print("  data: bytes length=", len(val))
                    break
                elif isinstance(val, str):
                    print("  data: text length=", len(val))
                    break
    else:
        print("No main resource found")

    # list subresources
    found = False
    for attr in ("subresources", "_subresources", "resources", "web_resources", "WebResources"):
        if hasattr(wa, attr):
            col = getattr(wa, attr)
            try:
                items = list(col) if col is not None else []
            except Exception:
                items = []
            if items:
                print(f"Found resources via {attr}: count={len(items)}")
                found = True
                for i, r in enumerate(items, 1):
                    url = maybe_attr(r, "url", "URL", "filename", "path", "src")
                    mime = maybe_attr(r, "mimeType", "MIMEType", "contentType")
                    size = None
                    for key in ("data", "data_bytes", "content", "html", "value"):
                        if hasattr(r, key):
                            v = getattr(r, key)
                            if isinstance(v, (bytes, bytearray)):
                                size = len(v); break
                            if isinstance(v, str):
                                size = len(v); break
                    print(f"  [{i}] url={url!r} mime={mime!r} data_len={size}")
    if not found:
        print("No subresources found via known attributes")

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python inspect_webarchive.py path/to/file.webarchive")
        sys.exit(1)
    p = Path(sys.argv[1])
    if not p.exists():
        print("File not found:", p)
        sys.exit(2)
    wa = WebArchive(str(p))
    probe(wa)
