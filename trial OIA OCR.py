import os
import subprocess
import sys

# TESSERACT_PATH = r"C:\Users\IlaBarshilia\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# INPUT_PDF = r"C:\Users\IlaBarshilia\Downloads\PDF Text Parser at OIA\CSI AP Aging as of 02.20.26.pdf"
# OUTPUT_PDF = r"C:\Users\IlaBarshilia\Downloads\PDF Text Parser at OIA\Output CSI AP Aging as of 02.20.26.pdf"

# # Inject Tesseract into PATH for this process only
# env = os.environ.copy()
# env["PATH"] = os.path.dirname(TESSERACT_PATH) + os.pathsep + env.get("PATH", "")

# # Run OCRmyPDF using the same Python interpreter
# subprocess.run(
#     [sys.executable, "-m", "ocrmypdf", "--force-ocr", INPUT_PDF, OUTPUT_PDF],
#     env=env,
#     check=True
# )

# print("✅ OCR completed successfully")

import pdfplumber
import pandas as pd
import re

# ============================================================
# CONFIGURATION
# ============================================================
PDF_PATH = r"C:\Users\IlaBarshilia\Downloads\PDF Text Parser at OIA\Output CSI AP Aging as of 02.20.26.pdf"

Y_TOL = 5.0   # tolerance for top/bottom (row grouping)

# Absolute column boundaries derived from your PDF
COLUMN_BOUNDS = [
    ("Vendor_Block",  20, 180),
    ("Due_Type",     180, 260),
    ("Amount",       260, 310),
    ("Current",      310, 360),
    ("Over_30",      380, 420),
    ("Over_60",      430, 470),
    ("Over_90",      480, 520),
    ("Over_120",     530, 580),
]

NUMERIC_COLS = [
    "Amount", "Current", "Over_30", "Over_60", "Over_90", "Over_120"
]

# ============================================================
# STEP 1: EXTRACT WORDS WITH COORDINATES
# ============================================================
words = []

with pdfplumber.open(PDF_PATH) as pdf:
    for page_no, page in enumerate(pdf.pages, start=1):
        for w in page.extract_words(
            use_text_flow=False,
            keep_blank_chars=True
        ):
            words.append({
                "page": page_no,
                "text": w["text"],
                "x0": w["x0"],
                "x1": w["x1"],
                "top": w["top"],
                "bottom": w["bottom"]
            })

df = pd.DataFrame(words)
df = df.sort_values(["page", "top", "x0"]).reset_index(drop=True)

# ============================================================
# STEP 2: GROUP WORDS INTO VISUAL ROWS (CLUSTER‑BASED)
# ============================================================
rows = []

for _, w in df.iterrows():
    placed = False
    for row in rows:
        if w["page"] != row[0]["page"]:
            continue

        # Compare against ANY word in the row (not just first)
        for rw in row:
            if (
                abs(w["top"] - rw["top"]) <= Y_TOL and
                abs(w["bottom"] - rw["bottom"]) <= Y_TOL
            ):
                row.append(w)
                placed = True
                break

        if placed:
            break

    if not placed:
        rows.append([w])

# ============================================================
# STEP 3: ASSIGN WORDS TO ABSOLUTE COLUMNS
# ============================================================
def assign_column(x_center):
    for col_name, x_min, x_max in COLUMN_BOUNDS:
        if x_min <= x_center < x_max:
            return col_name
    return None


final_rows = []

for row_id, row_words in enumerate(rows, start=1):
    row_data = {col: "" for col, _, _ in COLUMN_BOUNDS}
    row_data["page"] = row_words[0]["page"]
    row_data["row_id"] = row_id

    for w in row_words:
        x_center = (w["x0"] + w["x1"]) / 2
        col = assign_column(x_center)

        if col:
            if row_data[col]:
                row_data[col] += " " + w["text"]
            else:
                row_data[col] = w["text"]

    final_rows.append(row_data)

final_df = pd.DataFrame(final_rows)

# ============================================================
# STEP 4: NORMALIZE NUMBERS (AFTER ALIGNMENT)
# ============================================================
def normalize_number(val):
    if not val:
        return val

    v = val.strip()
    v = v.replace(" ", "")
    v = v.replace(",", ".")

    # fix common OCR artifacts seen in your PDF
    if re.fullmatch(r"9\.00", v):
        return "0.00"
    if re.fullmatch(r"\.00", v):
        return "0.00"

    return v


for col in NUMERIC_COLS:
    final_df[col] = final_df[col].apply(normalize_number)

# ============================================================
# FINAL CLEAN OUTPUT
# ============================================================
final_df.to_excel('Check_OCR_raw_DAta.xlsx')
print(final_df.head(30))

