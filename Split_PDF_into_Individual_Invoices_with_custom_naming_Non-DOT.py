import fitz
import pandas as pd
from datetime import datetime
import os
import re
import shutil

# ---------------------------------------------------------
# Minimal filename safety:
# Only replace characters that *must* be replaced on Windows.
# Do NOT touch commas, dots, or spaces inside names.
# ---------------------------------------------------------
def sanitize_filename_minimal(name):
    # Replace only illegal Windows characters: \ / : * ? " < > |
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    # Remove trailing space or dot (Windows does not allow)
    name = name.rstrip(" .")
    return name

# ---------------------------------------------------------
# Path-length safety only
# ---------------------------------------------------------
def safe_path(output_dir, filename):
    filename = sanitize_filename_minimal(filename)
    output_path = os.path.normpath(os.path.join(output_dir, filename))

    if len(output_path) > 240:  # Windows safe limit
        base, ext = os.path.splitext(filename)
        base = base[:100]  # truncate only if necessary
        filename = base + ext
        filename = sanitize_filename_minimal(filename)
        output_path = os.path.join(output_dir, filename)

    return output_path

# ---------------------------------------------------------
# Extract PDF structured text
# ---------------------------------------------------------
def extract_structured_text(pdf_path):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"❌ Error opening PDF: {e}")
        return pd.DataFrame()

    all_blocks = []
    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" in b:
                for line in b["lines"]:
                    line_text = " ".join([span["text"] for span in line["spans"]])
                    bbox = tuple(round(coord) for coord in line["bbox"])
                    all_blocks.append({
                        "page": page_num + 1,
                        "text": line_text,
                        "x0": bbox[0],
                        "y0": bbox[1],
                        "x1": bbox[2],
                        "y1": bbox[3]
                    })
    return pd.DataFrame(all_blocks).sort_values(by=["page", "y0", "x0"])

# ---------------------------------------------------------
# Positional helpers
# ---------------------------------------------------------
def find_text_right_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            candidates = df[
                (df["page"] == row["page"]) &
                (df["x0"] > row["x1"]) &
                (df["y0"] >= row["y0"] - tolerance) &
                (df["y1"] <= row["y1"] + tolerance)
            ]
            if not candidates.empty:
                result.append({"page": row["page"], "right_text": candidates.iloc[0]["text"]})
    return pd.DataFrame(result)

def find_text_left_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            candidates = df[
                (df["page"] == row["page"]) &
                (df["x1"] < row["x0"]) &
                (df["y0"] >= row["y0"] - tolerance) &
                (df["y1"] <= row["y1"] + tolerance)
            ]
            if not candidates.empty:
                result.append({"page": row["page"], "left_text": candidates.iloc[-1]["text"]})
    return pd.DataFrame(result)

# ---------------------------------------------------------
# Main Split Logic
# ---------------------------------------------------------
def split_pdf_with_custom_naming(pdf_path1, pdf_filename):
    try:
        pdf_path = os.path.join(pdf_path1, pdf_filename + ".pdf")
        temp_dir = "C:\\Temp\\split_invoices"
        os.makedirs(temp_dir, exist_ok=True)

        df = extract_structured_text(pdf_path)
        if df.empty:
            print("⚠️ No text extracted. Skipping split.")
            return

        invoice_df = find_text_right_of(df, "Invoice No")
        date_df = find_text_right_of(df, "Date")
        customer_df = find_text_left_of(df, "Job No")

        doc = fitz.open(pdf_path)
        saved_files = []

        current_invoice_number = "Unknown"
        current_date = "UnknownDate"
        current_customer_name = "UnknownCustomer"
        current_pages = []

        for page_num in range(len(doc)):
            page_index = page_num + 1
            invoice_number = invoice_df[invoice_df["page"] == page_index]["right_text"].values
            date = date_df[date_df["page"] == page_index]["right_text"].values
            customer_name = customer_df[customer_df["page"] == page_index]["left_text"].values

            if len(invoice_number) > 0:

                # Save previous packet
                if current_pages:
                    filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"
                    filename = sanitize_filename_minimal(filename)
                    output_path = safe_path(temp_dir, filename)

                    new_doc = fitz.open()
                    for p in current_pages:
                        new_doc.insert_pdf(doc, from_page=p, to_page=p)
                    try:
                        new_doc.save(output_path)
                    except Exception as e:
                        print(f"❌ Failed to save file: {e}")
                    new_doc.close()
                    saved_files.append(output_path)

                # Start new packet
                current_invoice_number = invoice_number[0]
                current_date = date[0] if len(date) > 0 else "UnknownDate"
                current_customer_name = customer_name[0] if len(customer_name) > 0 else "UnknownCustomer"
                current_pages = [page_num]

            else:
                current_pages.append(page_num)

        # Final invoice save
        if current_pages:
            filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"
            filename = sanitize_filename_minimal(filename)
            output_path = safe_path(temp_dir, filename)

            new_doc = fitz.open()
            for p in current_pages:
                new_doc.insert_pdf(doc, from_page=p, to_page=p)
            try:
                new_doc.save(output_path)
            except Exception as e:
                print(f"❌ Failed to save file: {e}")
            new_doc.close()
            saved_files.append(output_path)

        # Move to final directory
        final_dir = os.path.join(
            pdf_path1,
            f"split_invoices_of_{pdf_filename}_{datetime.today().strftime('%Y-%b-%d')}"
        )
        os.makedirs(final_dir, exist_ok=True)

        for file in saved_files:
            shutil.move(file, final_dir)

        print(f"\n✅ Saved {len(saved_files)} files in '{final_dir}'")

    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")

# ---------------------------------------------------------
# Loop
# ---------------------------------------------------------
while True:
    pdf_path1 = input("\nEnter the path to where your PDF file is kept: ").strip()
    pdf_filename = input("Enter the name of the PDF file (without .pdf): ").strip()

    pdf_path = os.path.join(pdf_path1, pdf_filename + ".pdf")
    if not os.path.exists(pdf_path):
        print(f"❌ File '{pdf_filename}.pdf' not found in '{pdf_path1}'.")
        break

    split_pdf_with_custom_naming(pdf_path1, pdf_filename)

    again = input("\nSplit another PDF? (Yes/No): ").strip().lower()
    if again != "yes":
        print("👋 Exiting.")
        break