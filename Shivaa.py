#!/usr/bin/env python
# coding: utf-8

# In[ ]:





# In[2]:


import pandas as pd
import numpy as np
import re
from dateutil import parser

# ---------------- CONFIG ----------------
file_path = r"C:\Users\IlaBarshilia\Downloads\Dummy Data Power BI formated for me.xlsx"
output_path = r"C:\Users\IlaBarshilia\Downloads\LDIProject_final_sheet1.xlsx"

cutoff_row = 10
header_row = 11
invoice_row = 13
sheets_to_process = ["Sheet1", "Sheet2", "Sheet3"]
all_data = []

print("📘 Starting LDI Final Extraction...\n")

# ================================================================
# FUNCTION — Detect Contract Number
# ================================================================
def detect_contract_number(df):
    contract_no = ""
    for row_idx in range(min(20, len(df))):
        row_text = " ".join(df.iloc[row_idx].astype(str).tolist())
        match = re.search(r"\bT\d{3,6}\b", row_text, re.IGNORECASE)
        if match:
            contract_no = match.group(0).upper()
            break
    return contract_no if contract_no else "N/A"

# ================================================================
# FUNCTION 1 — SIMPLE SHEETS (Sheet1 & Sheet2)
# ================================================================
def process_simple_sheet(sheet):
    print(f"\n🟢 Processing (simple layout): {sheet}")
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    contract_no = detect_contract_number(df)

    cutoff_dates = df.iloc[cutoff_row].fillna("").tolist()
    acme_dot_headers = df.iloc[header_row].fillna("").tolist()
    invoice_numbers = df.iloc[invoice_row].fillna("").tolist()

    pattern = r"(?i)Mobilization|Work Zone|Type|Pedestrian|Arrow|Message"
    data_start_row = df[
        df.astype(str).apply(lambda col: col.str.contains(pattern, case=False, na=False)).any(axis=1)
    ].index.min()
    data_start_row = int(data_start_row) if not pd.isna(data_start_row) else 15

    data = df.iloc[data_start_row:].reset_index(drop=True)
    data.columns = [f"Column_{i}" for i in range(len(data.columns))]

    # Detect price column
    price_col_idx = None
    for idx, val in enumerate(acme_dot_headers):
        if "price" in str(val).lower():
            price_col_idx = idx
            break

    clean_data = []
    skip_words = [
        "note", "notes", "retail", "subtotal", "estimate", "payment",
        "bond", "extras", "non contract", "total", "nan", "description",
        "jax", "svm", "pay app"
    ]

    for i in range(2, len(acme_dot_headers), 2):
        acme_col = f"Column_{i}"
        dot_col = f"Column_{i+1}"

        if "ACME" in str(acme_dot_headers[i]).upper() and "DOT" in str(acme_dot_headers[i+1]).upper():
            cutoff_date = cutoff_dates[i]
            invoice_no = invoice_numbers[i]

            for _, row in data.iterrows():
                desc = str(row["Column_0"]).strip().replace(";", "")
                if not desc or any(word in desc.lower() for word in skip_words):
                    continue

                code = str(row["Column_1"]).strip()
                acme_val = pd.to_numeric(row[acme_col], errors="coerce")
                dot_val = pd.to_numeric(row[dot_col], errors="coerce")

                acme_val = 0 if pd.isna(acme_val) else acme_val
                dot_val = 0 if pd.isna(dot_val) else dot_val

                price_val = 0
                if price_col_idx is not None and price_col_idx < len(data.columns):
                    price_val = pd.to_numeric(row[f"Column_{price_col_idx}"], errors="coerce")
                    price_val = 0 if pd.isna(price_val) else price_val

                # clean cutoff date → remove "00:00:00" everywhere
                if isinstance(cutoff_date, str):
                    cutoff_date = cutoff_date.replace("00:00:00", "").strip()

                clean_data.append({
                    "Source Sheet": sheet,
                    "Item Description": desc,
                    "Item Code": code if code.lower() != "nan" else "",
                    "Invoice No": invoice_no,
                    "Cutoff Date": cutoff_date,
                    "ACME Quantity": acme_val,
                    "DOT Quantity": dot_val,
                    "Price": price_val,
                    "Contract No": contract_no
                })

    print(f"✅ Extracted {len(clean_data)} rows from {sheet}")
    return pd.DataFrame(clean_data)

# ================================================================
# FUNCTION 2 — BLOCK LOGIC (Sheet3)
# ================================================================
def process_block_sheet(sheet):
    print(f"\n🧩 Processing (block layout): {sheet}")
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    contract_no = detect_contract_number(df)

    date_rx = r"(?:\d{1,2}/\d{1,2}/\d{4}|January|February|March|April|May|June|July|August|September|October|November|December)"
    date_row_candidates = df[df.astype(str).apply(lambda col: col.str.contains(date_rx, case=False, na=False, regex=True)).any(axis=1)].index.tolist()
    cutoff_row = date_row_candidates[0]
    header_row = cutoff_row + 1

    raw_cutoff = df.iloc[cutoff_row].fillna("").astype(str).tolist()
    cutoff_dates = [v.strip() for v in raw_cutoff]
    hdr = df.iloc[header_row].fillna("").astype(str).tolist()

    acme_dot_pairs = []
    for i in range(len(hdr) - 1):
        if "ACME" in str(hdr[i]).upper() and "DOT" in str(hdr[i+1]).upper():
            acme_dot_pairs.append((i, i+1))

    price_col_idx = None
    for idx, val in enumerate(hdr):
        if "price" in str(val).lower():
            price_col_idx = idx
            break

    invoice_rows = df[df.astype(str).apply(lambda col: col.str.contains("INVOICE", case=False, na=False)).any(axis=1)].index.tolist()
    skip_words = [
        "extras", "note", "notes", "subtotal", "estimate", "payment", "bond",
        "non contract", "jax", "svm", "pay app", "total", "retail", "description"
    ]

    def invoice_from_row_col(row_idx, col_idx):
        if row_idx >= len(df):
            return ""
        val = str(df.iat[row_idx, col_idx]) if col_idx < df.shape[1] else ""
        nums = re.findall(r"\b\d{5,7}\b", val)
        return ", ".join(sorted(set(nums), key=nums.index)) if nums else ""

    rows_out = []
    for idx, inv_row in enumerate(invoice_rows):
        start = inv_row + 1
        end = invoice_rows[idx + 1] if idx + 1 < len(invoice_rows) else len(df)
        block_df = df.iloc[start:end].fillna("").reset_index(drop=True)

        for acme_col, dot_col in acme_dot_pairs:
            inv_val = invoice_from_row_col(inv_row, acme_col)
            cutoff_val = cutoff_dates[acme_col] if acme_col < len(cutoff_dates) else ""

            if isinstance(cutoff_val, str):
                cutoff_val = cutoff_val.replace("00:00:00", "").strip()

            for ridx in range(block_df.shape[0]):
                desc = str(block_df.iat[ridx, 0]).strip().replace(";", "")
                if not desc or any(w in desc.lower() for w in skip_words):
                    continue

                code = str(block_df.iat[ridx, 1]).strip() if block_df.shape[1] > 1 else ""
                acme_val = pd.to_numeric(block_df.iat[ridx, acme_col], errors='coerce')
                dot_val = pd.to_numeric(block_df.iat[ridx, dot_col], errors='coerce')

                acme_val = 0 if pd.isna(acme_val) else acme_val
                dot_val = 0 if pd.isna(dot_val) else dot_val

                price_val = 0
                if price_col_idx is not None and price_col_idx < block_df.shape[1]:
                    price_val = pd.to_numeric(block_df.iat[ridx, price_col_idx], errors='coerce')
                    price_val = 0 if pd.isna(price_val) else price_val

                rows_out.append({
                    "Source Sheet": sheet,
                    "Item Description": desc,
                    "Item Code": code,
                    "Invoice No": inv_val,
                    "Cutoff Date": cutoff_val,
                    "ACME Quantity": acme_val,
                    "DOT Quantity": dot_val,
                    "Price": price_val,
                    "Contract No": contract_no
                })

    print(f"✅ Extracted {len(rows_out)} rows from {sheet}")
    return pd.DataFrame(rows_out)

# ================================================================
# MAIN LOOP
# ================================================================
for sheet in sheets_to_process:
    df_out = process_block_sheet(sheet) if sheet == "Sheet3" else process_simple_sheet(sheet)
    all_data.append(df_out)

# ================================================================
# COMBINE & SAVE
# ================================================================
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    final_df = final_df[final_df["Item Description"].notna()]
    final_df = final_df[~final_df["Item Description"].str.lower().isin(["nan", "description"])]

    # fill blanks in quantities with 0
    final_df["ACME Quantity"] = final_df["ACME Quantity"].fillna(0)
    final_df["DOT Quantity"] = final_df["DOT Quantity"].fillna(0)

    final_df.to_excel(output_path, index=False)
    print("\n🎯 LDI FINAL FILE CREATED SUCCESSFULLY (Cleaned & Correct Output)")
    print(f"💾 Saved to: {output_path}")
    print(f"📊 Total Rows: {len(final_df)}")
    display(final_df.head(25))
else:
    print("❌ No data extracted from any sheet.")


# In[3]:


import pandas as pd
import numpy as np
import re
from dateutil import parser

# ---------------- CONFIG ----------------
file_path = r"C:\Users\PoojaKale\Downloads\Dummy Data Power BI formated for me.xlsx"
output_path = r"C:\Users\PoojaKale\Downloads\LDI_Final_Output.xlsx"

cutoff_row = 10
header_row = 11
invoice_row = 13
sheets_to_process = ["Sheet1", "Sheet2", "Sheet3"]
all_data = []

print("📘 Starting LDI Final Extraction...\n")

# ================================================================
# FUNCTION — Detect Contract Number
# ================================================================
def detect_contract_number(df):
    contract_no = ""
    for row_idx in range(min(20, len(df))):
        row_text = " ".join(df.iloc[row_idx].astype(str).tolist())
        match = re.search(r"\bT\d{3,6}\b", row_text, re.IGNORECASE)
        if match:
            contract_no = match.group(0).upper()
            break
    return contract_no if contract_no else "N/A"

# ================================================================
# FUNCTION 1 — SIMPLE SHEETS (Sheet1 & Sheet2)
# ================================================================
def process_simple_sheet(sheet):
    print(f"\n🟢 Processing (simple layout): {sheet}")
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    contract_no = detect_contract_number(df)

    cutoff_dates = df.iloc[cutoff_row].fillna("").astype(str).str.replace("00:00:00", "").str.strip().tolist()
    acme_dot_headers = df.iloc[header_row].fillna("").tolist()
    invoice_numbers = df.iloc[invoice_row].fillna("").tolist()

    pattern = r"(?i)Mobilization|Work Zone|Type|Pedestrian|Arrow|Message"
    start_row = df[df.astype(str).apply(lambda col: col.str.contains(pattern, na=False, case=False)).any(axis=1)].index.min()
    data_start_row = 15 if pd.isna(start_row) else int(start_row)

    data = df.iloc[data_start_row:].reset_index(drop=True)
    data.columns = [f"Column_{i}" for i in range(len(data.columns))]

    # Detect Price col
    price_col_idx = None
    for idx, val in enumerate(acme_dot_headers):
        if "price" in str(val).lower():
            price_col_idx = idx
            break

    clean_data = []
    skip_words = ["note", "notes", "subtotal", "estimate", "payment", "bond",
                  "retail", "extras", "total", "description", "non contract",
                  "jax", "svm", "pay app", "nan"]

    for i in range(2, len(acme_dot_headers), 2):
        acme_col = f"Column_{i}"
        dot_col = f"Column_{i+1}"

        if "ACME" in str(acme_dot_headers[i]).upper() and "DOT" in str(acme_dot_headers[i+1]).upper():
            cutoff_date = cutoff_dates[i]
            invoice_no = invoice_numbers[i]

            for _, row in data.iterrows():
                desc = str(row["Column_0"]).strip().replace(";", "")
                if not desc or any(w in desc.lower() for w in skip_words):
                    continue

                code = str(row["Column_1"]).strip()
                acme_val = pd.to_numeric(row[acme_col], errors="coerce")
                dot_val = pd.to_numeric(row[dot_col], errors="coerce")

                acme_val = 0 if pd.isna(acme_val) else acme_val
                dot_val = 0 if pd.isna(dot_val) else dot_val

                # Price
                price_val = 0
                if price_col_idx is not None:
                    price_val = pd.to_numeric(row.get(f"Column_{price_col_idx}", 0), errors="coerce")
                    price_val = 0 if pd.isna(price_val) else price_val

                clean_data.append({
                    "Source Sheet": sheet,
                    "Item Description": desc,
                    "Item Code": code if code.lower() != "nan" else "",
                    "Invoice No": invoice_no,
                    "Cutoff Date": cutoff_date,
                    "ACME Quantity": acme_val,
                    "DOT Quantity": dot_val,
                    "Price": price_val,
                    "Contract No": contract_no
                })

    print(f"✅ Extracted {len(clean_data)} rows from {sheet}")
    return pd.DataFrame(clean_data)

# ================================================================
# FUNCTION 2 — BLOCK LOGIC (Sheet3 - Use your PERFECT invoice logic)
# ================================================================
def process_block_sheet(sheet):
    print(f"\n🧩 Processing Sheet3 (block layout with PERFECT invoice logic)...")
    df = pd.read_excel(file_path, sheet_name=sheet, header=None)
    contract_no = detect_contract_number(df)

    # detect cutoff row
    date_rx = r"(?:\d{1,2}/\d{1,2}/\d{4}|January|February|March|April|May|June|July|August|September|October|November|December)"
    cutoff_row = df[df.astype(str).apply(lambda col: col.str.contains(date_rx, na=False, regex=True)).any(axis=1)].index[0]
    header_row = cutoff_row + 1

    cutoff_dates = df.iloc[cutoff_row].fillna("").astype(str).str.replace("00:00:00", "").str.strip().tolist()
    hdr = df.iloc[header_row].fillna("").astype(str).tolist()

    acme_dot_pairs = []
    for i in range(len(hdr) - 1):
        if "ACME" in hdr[i].upper() and "DOT" in hdr[i+1].upper():
            acme_dot_pairs.append((i, i+1))

    invoice_rows = df[df.astype(str).apply(lambda col: col.str.contains("INVOICE", na=False, case=False)).any(axis=1)].index.tolist()

    skip_words = ["extras", "note", "subtotal", "estimate", "payment", "bond",
                  "non contract", "jax", "svm", "pay app", "total", "retail",
                  "description"]

    def invoice_from_row_col(row_idx, col_idx):
        """ PERFECT invoice logic (your working code) """
        val = str(df.iat[row_idx, col_idx]) if col_idx < df.shape[1] else ""
        val = re.sub(r"(?i)inv\s*#?:?", "", val)
        val = re.sub(r"[^0-9,/ ]", "", val).strip()
        parts = [p.strip() for p in re.split(r"[,/ ]+", val) if p.strip()]
        return parts[0] if parts else ""

    rows_out = []
    for idx, inv_row in enumerate(invoice_rows):
        start = inv_row + 1
        end = invoice_rows[idx+1] if idx+1 < len(invoice_rows) else len(df)

        block_df = df.iloc[start:end].fillna("").reset_index(drop=True)

        for acme_col, dot_col in acme_dot_pairs:
            inv_val = invoice_from_row_col(inv_row, acme_col)
            cutoff_val = cutoff_dates[acme_col] if acme_col < len(cutoff_dates) else ""

            for ridx in range(block_df.shape[0]):
                desc = str(block_df.iat[ridx, 0]).strip().replace(";", "")
                if not desc or any(w in desc.lower() for w in skip_words):
                    continue

                code = str(block_df.iat[ridx, 1]).strip()
                acme_val = pd.to_numeric(block_df.iat[ridx, acme_col], errors="coerce")
                dot_val = pd.to_numeric(block_df.iat[ridx, dot_col], errors="coerce")

                acme_val = 0 if pd.isna(acme_val) else acme_val
                dot_val = 0 if pd.isna(dot_val) else dot_val

                rows_out.append({
                    "Source Sheet": sheet,
                    "Item Description": desc,
                    "Item Code": code,
                    "Invoice No": inv_val,
                    "Cutoff Date": cutoff_val,
                    "ACME Quantity": acme_val,
                    "DOT Quantity": dot_val,
                    "Price": 0,
                    "Contract No": contract_no
                })

    print(f"✅ Extracted {len(rows_out)} rows from Sheet3")
    return pd.DataFrame(rows_out)

# ================================================================
# MAIN LOOP
# ================================================================
for sheet in sheets_to_process:
    if sheet == "Sheet3":
        df_out = process_block_sheet(sheet)
    else:
        df_out = process_simple_sheet(sheet)

    all_data.append(df_out)

# ================================================================
# COMBINE & SAVE → FINAL DELIVERY FILE
# ================================================================
final_df = pd.concat(all_data, ignore_index=True)

final_df["ACME Quantity"] = final_df["ACME Quantity"].fillna(0)
final_df["DOT Quantity"] = final_df["DOT Quantity"].fillna(0)

final_df.to_excel(output_path, index=False)
print("\n🎉 FINAL LDI FILE CREATED SUCCESSFULLY!")
print(f"💾 Saved to: {output_path}")
print(f"📊 Total Rows: {len(final_df)}")

display(final_df.head(25))


# In[ ]:




