import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime
import os
import re

pdf_path1 = input("Enter the path to where your PDF file is kept that needs to be split: ")
pdf_filename = input("Enter the name of the PDF file in the path shared above that you would like to split: ")
pdf_path = os.path.join(pdf_path1, pdf_filename)

if not os.path.exists(pdf_path):
    print(f"Error: The file '{pdf_filename}' does not exist in the path '{pdf_path1}'. Please check the path and filename.")
    input("Press Enter to exit...")
    exit(1)

def sanitize_filename(name):
    return re.sub(r'[\\\\/:*?"<>|]', '-', name)

def extract_structured_text(pdf_path):
    doc = fitz.open(pdf_path)
    all_blocks = []

    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        for b in blocks:
            if "lines" in b:
                for line in b["lines"]:
                    line_text = " ".join([span["text"] for span in line["spans"]])
                    bbox = line["bbox"]
                    bbox = tuple(round(coord) for coord in bbox)
                    all_blocks.append({
                        "page": page_num + 1,
                        "text": line_text,
                        "x0": bbox[0],
                        "y0": bbox[1],
                        "x1": bbox[2],
                        "y1": bbox[3]
                    })
    df = pd.DataFrame(all_blocks)
    return df.sort_values(by=["page", "y0", "x0"])

def find_text_right_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            x1_target = row["x1"]
            y0_target = row["y0"]
            y1_target = row["y1"]
            page = row["page"]

            candidates = df[
                (df["page"] == page) & 
                (df["x0"] > x1_target) &
                (df["y0"] >= y0_target - tolerance) &
                (df["y1"] <= y1_target + tolerance)
            ]

            if not candidates.empty:
                result.append({
                    "page": page,
                    "right_text": candidates.iloc[0]["text"]
                })
    return pd.DataFrame(result)

def find_text_left_of(df, target_text, tolerance=2):
    result = []
    for _, row in df.iterrows():
        if target_text in row["text"]:
            x0_target = row["x0"]
            y0_target = row["y0"]
            y1_target = row["y1"]
            page = row["page"]

            candidates = df[
                (df["page"] == page) &
                (df["x1"] < x0_target) &
                (df["y0"] >= y0_target - tolerance) &
                (df["y1"] <= y1_target + tolerance)
            ]

            if not candidates.empty:
                result.append({
                    "page": page,
                    "left_text": candidates.iloc[-1]["text"]
                })
    return pd.DataFrame(result)

def split_pdf_with_custom_naming(pdf_path):
    output_dir = os.path.join(pdf_path1, f"split_invoices_of_{pdf_filename}_{datetime.today().strftime('%Y-%b-%d')}")
    os.makedirs(output_dir, exist_ok=True)

    df = extract_structured_text(os.path.join(pdf_path1, pdf_filename + ".pdf"))

    invoice_df = find_text_right_of(df, "Invoice No", tolerance=2)
    date_df = find_text_right_of(df, "Date", tolerance=2)
    customer_df = find_text_left_of(df, "Job No", tolerance=2)

    doc = fitz.open(os.path.join(pdf_path1, pdf_filename + ".pdf"))
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
            # Save previous invoice if pages exist
            if current_pages:
                filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"
                filename = sanitize_filename(filename)
                output_path = os.path.join(output_dir, filename)

                new_doc = fitz.open()
                for p in current_pages:
                    new_doc.insert_pdf(doc, from_page=p, to_page=p)
                new_doc.save(output_path)
                new_doc.close()
                saved_files.append(output_path)

            # Start new invoice group
            current_invoice_number = invoice_number[0]
            current_date = date[0] if len(date) > 0 else "UnknownDate"
            current_customer_name = customer_name[0] if len(customer_name) > 0 else "UnknownCustomer"
            current_pages = [page_num]
        else:
            # Continue current invoice group
            current_pages.append(page_num)

    # Save the last invoice group
    if current_pages:
        filename = f"{current_customer_name} - Inv {current_invoice_number}, {current_date}.pdf"
        filename = sanitize_filename(filename)
        output_path = os.path.join(output_dir, filename)

        new_doc = fitz.open()
        for p in current_pages:
            new_doc.insert_pdf(doc, from_page=p, to_page=p)
        new_doc.save(output_path)
        new_doc.close()
        saved_files.append(output_path)

    
    print(f"\n✅ Saved {len(saved_files)} files in '{output_dir}'")
    input("Press Enter to exit...")
    return saved_files


split_pdf_with_custom_naming(pdf_path)