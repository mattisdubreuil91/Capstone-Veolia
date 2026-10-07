"""Extract statement tables from the Veolia PDFs into a CSV (label, values, document, page).

Usage: python extract_statements.py <pdf_dir> <out_csv>
Each row keeps the PDF page it came from so every number in the Excel can be traced.
"""
import csv, re, subprocess, sys, pathlib

NUM = r"\(?-?\d{1,3}(?:,\d{3})*(?:\.\d+)?\)?%?|-"
TAIL = re.compile(rf"^(?P<label>.*?\S)\s+(?:(?:Note|Notes)\s+[\d.]+(?:\s*&\s*Note\s+[\d.]+)?\s+|[\d.]+\s*&\s*Note\s*[\d.]+\s+)?(?P<nums>(?:{NUM})(?:\s+(?:{NUM}))+)\s*$")

# (document key, file prefix, pages, section name, number of value columns, column labels)
JOBS = [
    ("FY2025 statements", "ebd59084", [3], "Balance sheet - assets", ["31/12/2024", "31/12/2025"]),
    ("FY2025 statements", "ebd59084", [4], "Balance sheet - equity and liabilities", ["31/12/2024", "31/12/2025"]),
    ("FY2025 statements", "ebd59084", [5], "Income statement", ["FY2024", "FY2025"]),
    ("FY2025 statements", "ebd59084", [7, 8], "Cash flow statement", ["FY2024", "FY2025"]),
    ("H1-2026 report", "4834a434", [34], "Balance sheet - assets", ["31/12/2025", "30/06/2026"]),
    ("H1-2026 report", "4834a434", [35], "Balance sheet - equity and liabilities", ["31/12/2025", "30/06/2026"]),
    ("H1-2026 report", "4834a434", [36], "Income statement", ["H1 2025", "H1 2026"]),
    ("H1-2026 report", "4834a434", [38, 39], "Cash flow statement", ["H1 2025", "H1 2026"]),
    ("H1-2026 report", "4834a434", [6], "Key figures", ["30/06/2026", "31/12/2025", "30/06/2025", "31/12/2024"]),
    ("H1-2026 report", "4834a434", [26], "Net debt bridge", ["H1 2025", "H1 2026"]),
    ("H1-2024 report", "0a2e10f6", [28], "Net debt bridge", ["H1 2023", "H1 2024"]),
    ("H1-2024 report", "0a2e10f6", [6], "Key figures", ["30/06/2024", "31/12/2023", "30/06/2023", "31/12/2022"]),
]

def pages(pdf):
    out = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, check=True).stdout
    return out.split("\f")

def to_num(t):
    t = t.strip()
    if t == "-":
        return 0.0
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()").replace(",", "").rstrip("%")
    v = float(t)
    return -v if neg else v

def main(pdf_dir, out_csv):
    pdfs = {p.name[:8]: p for p in pathlib.Path(pdf_dir).glob("*.pdf")}
    cache, rows = {}, []
    for doc, key, pgs, section, cols in JOBS:
        if key not in cache:
            cache[key] = pages(pdfs[key])
        for pg in pgs:
            prev = ""
            for line in cache[key][pg - 1].splitlines():
                m = TAIL.match(line.strip())
                if not m:
                    txt = line.strip()
                    if txt and not re.search(r"\d{2}", txt):
                        prev = txt
                    elif txt:
                        prev = ""
                    continue
                nums = m.group("nums").split()
                if len(nums) < len(cols):
                    continue
                vals = [to_num(x) for x in nums[-len(cols):]]
                label = re.sub(r"\s{2,}", " ", m.group("label")).strip(" .:")
                label = re.sub(r"\s*\(\d\)$", "", label)
                label = label.replace("Interest on right of use (IFRS", "Interest on right of use (IFRS 16)")
                if (label[:1].islower() or label.startswith("(")) and prev:
                    label = (prev + " " + re.sub(r"^\(\d\)\s*", "", label)).strip()
                prev = ""
                if re.fullmatch(r"[\d\s.,/()-]*", label) or "VEOLIA" in label:
                    continue
                rows.append([doc, section, pg, label] + [""] * 0 + list(zip(cols, vals)))
    with open(out_csv, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["document", "section", "pdf_page", "label", "period", "value"])
        for r in rows:
            for col, v in r[4:]:
                w.writerow(r[:4] + [col, v])
    print(len(rows), "lines extracted")

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
