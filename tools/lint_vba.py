"""Flag VBA constructs the LibreOffice harness cannot parse (they are valid in Excel).
usage: python3 tools/lint_vba.py"""
import glob, re, sys
bad = 0
for p in sorted(glob.glob("src/*.bas")):
    for n, line in enumerate(open(p, encoding="ascii"), 1):
        code = line.split("'")[0]
        # single-line If whose Then and Else parts are both Sub calls with arguments
        m = re.match(r"\s*If .+ Then\s+([A-Za-z_][\w.]*)\s+[^=]+\s+Else\s+([A-Za-z_][\w.]*)\s+[^=]+$", code.strip("\r\n"))
        if m:
            print(f"{p}:{n}: one-line If with Sub calls in both branches -> use a block If"); bad += 1
        if re.search(r"\b[A-Za-z_][\w.]*\s+\([^()]*\)\s*[\\/*+-].*,", code) and not re.search(r"=|\bIf\b|\bThen\b", code.split(",")[0]):
            print(f"{p}:{n}: Sub call whose first argument starts with '(' -> compute it first"); bad += 1
        if re.search(r"Function \w+\(.*\) As \w+\(\)", code):
            print(f"{p}:{n}: function returning a typed array -> return Variant"); bad += 1
        if re.search(r"\bm\w+\((nm|key)\)", code) and "Item" not in code and "Collection" in open(p).read():
            pass
sys.exit(1 if bad else 0)
