import fitz
import pandas as pd
from datetime import datetime
import os
import re
import shutil

# ---------------------------------------------------------
# Minimal filename safety: only what Windows requires to avoid failure
# ---------------------------------------------------------
_ILLEGAL_CHARS_RE = re.compile(r'[\\/:*?"<>|]')  # only characters Windows forbids in filenames
_RESERVED_NAMES = {
    "CON","PRN","AUX","NUL",
    *(f"COM{i}" for i in range(1,10)),
    *(f"LPT{i}" for i in range(1,10)),
}

def minimal_safe_filename(name: str) -> str:
    """
    Apply only minimal changes needed so Windows can save the file:
      - Replace illegal characters \ / : * ? " < > |
      - Strip trailing space or period (Windows disallows trailing . or space)
      - Avoid reserved device names by appending underscore
    No other changes (commas, periods, spaces inside are kept).
    """
    # Replace only illegal characters
    safe = _ILLEGAL_CHARS_RE.sub("-", name)

    # Windows forbids names that end with a dot or space
    safe = safe.rstrip(" .")

    # Avoid reserved device names (case-insensitive) for the basename without extension
    base, ext = os.path.splitext(safe)
    if base.upper() in _RESERVED_NAMES:
        base = base + "_"
    safe = base + ext

    return safe

def safe_path(output_dir, filename):
    """
    Keep text intact except minimal rules above. If full path too long, truncate filename.
    """
    filename = minimal_safe_filename(filename)
    output_path = os.path.normpath(os.path.join(output_dir, filename))
    # Conservative limit for classic Windows APIs
    if len(output_path) > 240:
        base, ext = os.path.splitext(filename)
        # Truncate base only if required
        max_base_len = max(1, 240 - len(output_dir) - len(ext) - 10)  # rough allowance for separators
        base = base[:max_base_len]
        filename = base + ext
        filename = minimal_safe_filename(filename)  # re-apply minimal rules after truncation
        output_path = os.path.join(output_dir, filename)
    return output_path

# ---------------------------------------------------------
# Extract structured text from PDF using PyMuPDF
# ---------------------------------------------------------
def extract_structured_text(pdf_path):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"❌ Error opening PDF: {e}")
        return pd.DataFrame()

    rows = []
    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict").get("blocks", [])
        for b in blocks:
            if "lines" not in b:
                continue
            for line in b["lines"]:
                text = " ".join([s["text"] for s in line.get("spans", [])]).strip()
                if not text:
                    continue
                x0, y0, x1, y1 = [round(v) for v in line["bbox"]]
                rows.append({
                    "page": page_num + 1,
                    "text": text,
                    "x0": x0, "y0": y0,
                    "x1": x1, "y1": y1
                })

    return pd.DataFrame(rows).sort_values(by=["page", "y0", "x0"])

# ---------------------------------------------------------
# Positional helpers
# ---------------------------------------------------------
def _same_line(row, target, tolerance=2):
    return (
        row["y0"] >= target["y0"] - tolerance and
        row["y1"] <= target["y1"] + tolerance
    )

def find_text_right_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            candidates = df[
                (df["page"] == row["page"]) &
                (df["x0"] >= row["x1"] - 2) &  # allow tiny overlap
                df.apply(lambda r: _same_line(r, row, tolerance), axis=1)
            ].sort_values(by=["x0"])

            if not candidates.empty:
                result.append({
                    "page": row["page"],
                    "right_text": candidates.iloc[0]["text"]
                })

    return pd.DataFrame(result)

def find_text_left_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            candidates = df[
                (df["page"] == row["page"]) &
                (df["x1"] <= row["x0"] + 2) &
                df.apply(lambda r: _same_line(r, row, tolerance), axis=1)
            ].sort_values(by=["x1"])

            if not candidates.empty:
                result.append({
                    "page": row["page"],
                    "left_text": candidates.iloc[-1]["text"]
                })

    return pd.DataFrame(result)

# ---------------------------------------------------------
# REF extraction: use find_text_right_of for literal "REF:"
# New rules added:
#  - Token is exactly 5 chars: first is a letter, next 4 are alphanumeric
#  - The 4 trailing chars MUST contain at least one digit (reject pure letters)
#  - Treat space OR '/' as the delimiter after the 5-char token
# ---------------------------------------------------------
def extract_ref_codes(df):
    # First char letter; next 4 alnum AND must include at least one digit
    core_pattern = re.compile(r'^[A-Za-z][A-Za-z0-9]{4}$')
    has_digit = re.compile(r'\d')

    ref_raw = find_text_right_of(df, "REF:")
    if ref_raw.empty:
        return pd.DataFrame(columns=["page", "ref_code"])

    results = []
    for _, r in ref_raw.iterrows():
        text = r["right_text"].strip()
        if not text:
            continue

        # Split on space or slash (first token may be followed by whitespace or '/')
        # This handles 'A1B2C / something' or 'A1B2C something'
        first_part = re.split(r'[\/\s]+', text, maxsplit=1)[0]

        # Check exactly 5 chars, pattern, and presence of at least one digit among last 4
        if len(first_part) == 5 and core_pattern.match(first_part) and has_digit.search(first_part[1:]):
            results.append({"page": r["page"], "ref_code": first_part})

    return pd.DataFrame(results)

# ---------------------------------------------------------
# Main logic: split by invoice; minimal filename rules only
# ---------------------------------------------------------
def split_pdf_with_custom_naming(pdf_path1, pdf_filename):
    try:
        pdf_path = os.path.join(pdf_path1, pdf_filename + ".pdf")
        temp_dir = r"C:\Temp\split_invoices"
        os.makedirs(temp_dir, exist_ok=True)

        df = extract_structured_text(pdf_path)
        if df.empty:
            print("⚠️ No text extracted, skipping.")
            return

        invoice_df = find_text_right_of(df, "Invoice No")
        date_df = find_text_right_of(df, "Date")
        customer_df = find_text_left_of(df, "Job No")
        ref_df = extract_ref_codes(df)

        doc = fitz.open(pdf_path)
        saved_files = []

        current_invoice_number = "Unknown"
        current_date = "UnknownDate"
        current_customer_name = "UnknownCustomer"
        current_ref_code = None
        current_pages = []

        for page_num in range(len(doc)):
            page_index = page_num + 1

            invoice_val = invoice_df[invoice_df["page"] == page_index]["right_text"].values
            date_val = date_df[date_df["page"] == page_index]["right_text"].values
            cust_val = customer_df[customer_df["page"] == page_index]["left_text"].values
            ref_val = ref_df[ref_df["page"] == page_index]["ref_code"].values

            # New invoice boundary detected
            if len(invoice_val) > 0:

                # Save previous packet (if any)
                if current_pages:
                    if current_ref_code:
                        filename = f"{current_customer_name} - {current_ref_code} - Inv {current_invoice_number}, {current_date}.pdf"
                    else:
                        filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"

                    output_path = safe_path(temp_dir, filename)

                    new_doc = fitz.open()
                    for p in current_pages:
                        new_doc.insert_pdf(doc, from_page=p, to_page=p)
                    try:
                        new_doc.save(output_path)
                    except OSError as e:
                        print(f"❌ Save failed for '{filename}': {e}")
                    finally:
                        new_doc.close()

                    if os.path.exists(output_path):
                        saved_files.append(output_path)

                # Start new packet
                current_invoice_number = invoice_val[0].strip()
                current_date = date_val[0].strip() if len(date_val) else "UnknownDate"
                current_customer_name = cust_val[0].strip() if len(cust_val) else "UnknownCustomer"
                current_ref_code = ref_val[0].strip() if len(ref_val) else None

                current_pages = [page_num]

            else:
                # Continuation page
                if current_ref_code is None and len(ref_val) > 0:
                    current_ref_code = ref_val[0].strip()
                current_pages.append(page_num)

        # Save final packet
        if current_pages:
            if current_ref_code:
                filename = f"{current_customer_name} - {current_ref_code} - Inv {current_invoice_number}, {current_date}.pdf"
            else:
                filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"

            output_path = safe_path(temp_dir, filename)

            new_doc = fitz.open()
            for p in current_pages:
                new_doc.insert_pdf(doc, from_page=p, to_page=p)
            try:
                new_doc.save(output_path)
            except OSError as e:
                print(f"❌ Save failed for '{filename}': {e}")
            finally:
                new_doc.close()

            if os.path.exists(output_path):
                saved_files.append(output_path)

        # Move created files
        final_dir = os.path.join(
            pdf_path1,
            f"split_invoices_of_{pdf_filename}_{datetime.today().strftime('%Y-%b-%d')}"
        )
        os.makedirs(final_dir, exist_ok=True)
        for file in saved_files:
            try:
                shutil.move(file, final_dir)
            except Exception as e:
                print(f"⚠️ Move failed for '{file}': {e}")

        print(f"\n✅ Saved {len(saved_files)} files in '{final_dir}'")

    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")

# ---------------------------------------------------------
# Loop
# ---------------------------------------------------------
if __name__ == "__main__":
    while True:
        pdf_path1 = input("\nEnter PDF folder path: ").strip()
        pdf_filename = input("Enter PDF filename (without .pdf): ").strip()

        pdf_path = os.path.join(pdf_path1, pdf_filename + ".pdf")
        if not os.path.exists(pdf_path):
            print("❌ File not found.")
            break

        split_pdf_with_custom_naming(pdf_path1, pdf_filename)

        again = input("\nSplit another PDF? (Yes/No): ").strip().lower()
        if again != "yes":
            print("👋 Exiting.")
            break