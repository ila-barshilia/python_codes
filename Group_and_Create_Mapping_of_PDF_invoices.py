import sys
import traceback

def main():
    import os
    import pandas as pd
    import fitz
    from collections import defaultdict
    import shutil
    import re
    import numpy as np

    # ====================
    # Configuration
    # ====================
    # ====================
    # Configuration
    # ====================
    input_folder = input("Enter the path to where your PDF files are kept that you want to email to customers: ").strip()

    # Path for output container
    base_output_folder = os.path.join(input_folder, "Grouped_PDFs")

    # If Grouped_PDFs already exists → clear it completely
    # If Grouped_PDFs exists, clear its contents but keep the folder
    if os.path.exists(base_output_folder):
        for entry in os.scandir(base_output_folder):
            path = entry.path
            try:
                if entry.is_file() or entry.is_symlink():
                    os.remove(path)
                elif entry.is_dir():
                    shutil.rmtree(path)
            except Exception as e:
                print(f"⚠️ Could not delete {path}: {e}")

    # Recreate fresh output folder
    os.makedirs(base_output_folder, exist_ok=True)

    # Special folders (created fresh each run)
    no_email_folder = os.path.join(base_output_folder, "NoEmailFound")
    os.makedirs(no_email_folder, exist_ok=True)

    do_not_email_folder = os.path.join(base_output_folder, "DoNotEmail")
    os.makedirs(do_not_email_folder, exist_ok=True)


    # --- Define sanitizer early so it exists before use in _make_folder_name ---
    def safe_folder_name(name: str) -> str:
        if not isinstance(name, str):
            name = str(name or "")
        name = re.sub(r'[<>:"/\\|?*]', '_', name)  # Windows-invalid chars
        name = re.sub(r'\s+', ' ', name).strip(' .')
        # Keep folder names short for OneDrive; 240 keeps plenty of headroom
        return name[:240] if len(name) > 240 else name

    # --- Safe move helper with collision handling ---
    def move_with_collision(src_path: str, dst_dir: str, dst_filename: str) -> str:
        """
        Moves src_path into dst_dir/dst_filename.
        If a file with that name exists, appends ' (1)', ' (2)', ... before the extension.
        Returns the final destination path.
        """
        os.makedirs(dst_dir, exist_ok=True)
        root, ext = os.path.splitext(dst_filename)
        candidate = os.path.join(dst_dir, dst_filename)
        if not os.path.exists(candidate):
            shutil.move(src_path, candidate)
            return candidate

        # Resolve collisions by appending (n)
        n = 1
        while True:
            alt_name = f"{root} ({n}){ext}"
            candidate = os.path.join(dst_dir, alt_name)
            if not os.path.exists(candidate):
                shutil.move(src_path, candidate)
                return candidate
            n += 1

    # Store all PDFs' structured data
    all_dfs = []

    # ===========================
    # Extract text with coordinates
    # ===========================
    def extract_structured_text(pdf_path):
        doc = fitz.open(pdf_path)
        all_blocks = []

        for page_num, page in enumerate(doc):
            blocks = page.get_text("dict")["blocks"]
            for b in blocks:
                if "lines" in b:
                    for line in b["lines"]:
                        for span in line["spans"]:
                            bbox = tuple(round(coord) for coord in span["bbox"])
                            all_blocks.append({
                                "page": page_num + 1,
                                "text": span["text"],
                                "x0": bbox[0],
                                "y0": bbox[1],
                                "x1": bbox[2],
                                "y1": bbox[3]
                            })

        df = pd.DataFrame(all_blocks)
        return df.sort_values(by=["page", "y0", "x0"])

    # ===========================
    # Group by dynamic columns
    # ===========================
    def group_by_dynamic_columns(df, y_tolerance=2, x_bin_width=20):
        df = df.sort_values(by=["page", "y0", "x0"]).reset_index(drop=True)
        structured_rows = []

        for page in df["page"].unique():
            page_df = df[df["page"] == page].copy()
            rows = []

            for _, row in page_df.iterrows():
                y0 = row["y0"]
                matched = False

                for r in rows:
                    if abs(r["y0"] - y0) <= y_tolerance:
                        r["cells"].append(row)
                        matched = True
                        break

                if not matched:
                    rows.append({"y0": y0, "cells": [row]})

            for r in rows:
                bins = defaultdict(list)
                for cell in r["cells"]:
                    bin_index = int(cell["x0"] // x_bin_width)
                    bins[bin_index].append(cell["text"])

                sorted_bins = [bins[i] for i in sorted(bins)]
                structured_row = [" ".join(bin) for bin in sorted_bins]
                structured_rows.append(structured_row)

        return pd.DataFrame(structured_rows)

    # ===========================
    # PROCESS ALL PDFs
    # ===========================
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)

            df = extract_structured_text(pdf_path)
            df2 = group_by_dynamic_columns(df)

            df2["source_pdf"] = filename
            all_dfs.append(df2)

    if not all_dfs:
        print("No PDFs found. Nothing to do.")
        return

    final_df = pd.concat(all_dfs, ignore_index=True)

    # Extract InvoiceDate from filename (MM-DD-YY as in "02-15-26")
    final_df['InvoiceDate'] = final_df['source_pdf'].str.extract(r'(\d{2}-\d{2}-\d{2})', expand=False).fillna("")

    # ===========================
    # Extract CustomerName and DOT from source_pdf (your rules)
    # Always " - Inv" exists.
    # If only one " - " before Inv -> NO DOT (customer only)
    # Else, token immediately before Inv is DOT iff:
    #   - exactly 5 chars
    #   - first letter, next 4 alnum
    #   - last 4 contain at least one digit
    # ===========================
    DOT_CORE_RE = re.compile(r'^[A-Za-z][A-Za-z0-9]{4}$')

    def is_dot_code(token: str) -> bool:
        if not isinstance(token, str):
            return False
        t = token.strip()
        if len(t) != 5:
            return False
        if not DOT_CORE_RE.match(t):
            return False
        return any(ch.isdigit() for ch in t[1:])

    base_names = final_df['source_pdf'].str.replace(r'\.pdf$', '', case=False, regex=True)

    def parse_customer_and_dot(fname: str):
        base = re.sub(r'\.pdf$', '', str(fname or ''), flags=re.IGNORECASE).strip()
        m = re.search(r'\s-\s*Inv\b', base, flags=re.IGNORECASE)
        if not m:
            # fallback (shouldn't happen per your input guarantee)
            return base, ""
        left = base[:m.start()]  # before " - Inv"
        if " - " not in left:
            return left.strip(), ""
        parts = [p.strip() for p in left.split(" - ") if p.strip()]
        if len(parts) == 1:
            return parts[0], ""
        candidate = parts[-1]
        if is_dot_code(candidate):
            customer = " - ".join(parts[:-1]).strip()
            return customer, candidate
        else:
            return left.strip(), ""

    parsed = final_df['source_pdf'].apply(parse_customer_and_dot)
    final_df['CustomerName'] = parsed.apply(lambda x: x[0])
    final_df['DOT'] = parsed.apply(lambda x: x[1])

    # ===========================
    # Extract right columns for DoNotEmail
    # ===========================
    def extract_right_concat(df: pd.DataFrame, keywords=("Job No", "P.O. #"),
                             in_column: str | None = None, sep: str = " ", case_insensitive: bool = True) -> pd.Series:
        cols = list(df.columns)
        fourth_last_idx = len(cols) - 4
        kw_set = {k.lower() if case_insensitive else k for k in keywords}

        def _match(cell) -> bool:
            if cell is None or (isinstance(cell, float) and np.isnan(cell)):
                return False
            s = str(cell)
            s_cmp = s.lower() if case_insensitive else s
            return any(k in s_cmp for k in kw_set)

        def _concat_row(row: pd.Series):
            if in_column is not None:
                if in_column not in df.columns:
                    raise ValueError(f"`in_column` '{in_column}' not in DataFrame.")
                pos = cols.index(in_column)
                if not _match(row[in_column]):
                    return np.nan
            else:
                pos = None
                for j, c in enumerate(cols):
                    if _match(row[c]):
                        pos = j
                        break
                if pos is None:
                    return np.nan

            start = pos + 1
            end = fourth_last_idx
            if start > end:
                return np.nan

            vals = row.iloc[start:end+1].tolist()
            cleaned = [str(v).strip() for v in vals if v and not pd.isna(v)]
            return sep.join(cleaned) if cleaned else np.nan

        return df.apply(_concat_row, axis=1)

    final_df["DoNotEmail"] = extract_right_concat(final_df, keywords=("Job No", "P.O.", "#"))

    final_df2 = final_df[['source_pdf', 'DoNotEmail']]
    final_df2 = final_df2[final_df2["DoNotEmail"].str.contains('Pay App|Adj Inv', case=False, na=False)]
    final_df2 = final_df2.drop_duplicates(subset=['source_pdf'], keep='first')

    # ===========================
    # Extract relevant rows for Email block
    # ===========================
    fd = final_df.reset_index(drop=True)

    keyword = "Email"
    columns_to_check = [0, 1, 2]

    matches = fd.index[
        fd[columns_to_check].astype(str).apply(
            lambda row: any(keyword.lower() in str(val).lower() for val in row),
            axis=1
        )
    ]

    rows_to_include = []
    for pos in matches:
        end_pos = min(pos + 3, len(fd))
        rows_to_include.extend(range(pos, end_pos))

    rows_to_include = sorted(set(rows_to_include))
    df_filtered = fd.iloc[rows_to_include].reset_index(drop=True)

    df_filtered[0] = df_filtered[0].astype(str).str.replace("\u00A0", " ").str.strip()
    df_filtered = df_filtered[(df_filtered[0] != 'Customer Phone') & (df_filtered[0] != 'Customer Fax')]
    df_filtered.drop(columns="DoNotEmail", inplace=True)

    # ===========================
    # Combine email text across rows
    # ===========================
    def combine_email_per_pdf_blocks(df_filtered: pd.DataFrame) -> pd.DataFrame:
        if 'source_pdf' not in df_filtered.columns:
            raise ValueError("Expected 'source_pdf' in df_filtered.")
        src_idx = df_filtered.columns.get_loc("source_pdf")
        end_col = src_idx - 1

        def norm(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return ""
            return str(x).replace("\u00A0", " ").strip()

        search_cols = [0, 1, 2]
        markers = {"email", "none", "nan", ""}

        out_rows = []

        for pdf, grp in df_filtered.groupby("source_pdf", sort=False):
            email_pos = None
            first_match_row = None

            for ridx, row in grp.iterrows():
                for c in search_cols:
                    if c > end_col or c >= len(row):
                        continue
                    if "email" in norm(row.iloc[c]).lower():
                        email_pos = c
                        first_match_row = ridx
                        break
                if email_pos is not None:
                    break

            if email_pos is None or end_col < 0:
                out_rows.append({"source_pdf": pdf, "Email": ""})
                continue

            pieces = []
            for ridx, row in grp.iterrows():
                start_col = (email_pos + 1) if ridx == first_match_row else 0
                start_col = max(0, min(start_col, end_col))
                for col_idx in range(start_col, end_col + 1):
                    val = norm(row.iloc[col_idx])
                    if not val or val.lower() in markers:
                        continue
                    pieces.append(val)

            combined = "".join(pieces)
            out_rows.append({"source_pdf": pdf, "Email": combined})

        return pd.DataFrame(out_rows, columns=["source_pdf", "Email"])

    df_email_by_pdf = combine_email_per_pdf_blocks(df_filtered)

    # ============================================
    # Add RawEmail BEFORE cleaning
    # ============================================
    df_email_by_pdf["RawEmail"] = df_email_by_pdf["Email"].fillna("").astype(str)

    # Replace df_filtered for downstream
    df_filtered = df_email_by_pdf[['source_pdf', 'Email', 'RawEmail']].copy()
    df_filtered = df_filtered.groupby('source_pdf', as_index=False).agg({
        'Email': lambda x: ''.join(x),
        'RawEmail': 'first'
    })

    df_filtered['Email'] = df_filtered['Email'].str.replace(r'\*+', '', regex=True).str.strip()

    # ===========================
    # Clean Email (comments, cc, slashes)
    # ===========================
    def extract_comments_and_clean(email_str):
        s = email_str or ""
        comments = []

        if re.search(r'\bUS\s*Mail\b', s, flags=re.IGNORECASE):
            comments.append("US Mail")
        if re.search(r'\bUSMail\b', s, flags=re.IGNORECASE):
            comments.append("USMail")
        if re.search(r'\bread\s+receipt\b', s, flags=re.IGNORECASE):
            comments.append("read receipt")
        # Tidy some odd encodings of '& PM' seen in OCR/exports:
        s = re.sub(r'(?i)\s*/\s*US\s*Mail', '', s)
        s = re.sub(r'(?i)\s*/\s*USMail', '', s)
        s = re.sub(r'(?i)\s*\(\s*read\s+receipt\s*\)\s*', '', s)
        s = re.sub(r'(?i)\s*read\s+receipt\s*', '', s)
        s = re.sub(r'(?i)\s*&\s*PM\b', '', s)

        s = re.sub(r'\s+', ' ', s)
        s = re.sub(r'\s*,\s*', ', ', s)
        s = re.sub(r'(,\s*){2,}', ', ', s)
        s = s.strip(" ,")

        return s, "; ".join(dict.fromkeys(comments))  # de-dupe while preserving order

    clean_cols = df_filtered['Email'].apply(extract_comments_and_clean)
    df_filtered['Email'] = clean_cols.apply(lambda x: x[0])
    df_filtered['Comment'] = clean_cols.apply(lambda x: x[1])

    # Slashes and cc normalization
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'(?i)/\s*cc\s*([A-Za-z0-9._%+-]+)', r';\1', regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'(?i)/\s*cc\b', r';', regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'\s*/cc\s*', ';', regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'/', ';', regex=True)

    # ===========================
    # Normalize domains (complete tokens that end with '@')
    # ===========================
    def normalize_domains(email_str: str) -> str:
        """
        - Find the first valid domain from any complete email (e.g., '@example.com').
        - For tokens:
            * 'local'            -> 'local@domain'
            * 'local@'           -> 'local@domain'
            * 'local@bad'        -> 'local@domain' (if 'bad' isn't a valid domain)
            * 'local@valid.com'  -> keep
        - Preserve order and de-duplicate.
        """
        if not isinstance(email_str, str) or not email_str.strip():
            return email_str

        s = email_str.strip()
        # unify separators
        s = s.replace(',', ';')
        parts = [p.strip() for p in s.split(';') if p.strip()]

        if not parts:
            return ""

        local_re = re.compile(r'^[A-Za-z0-9._%+-]+$')
        domain_re = re.compile(r'^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$')  # simple, robust domain pattern

        # Discover reference domain from the first valid full email
        ref_domain = None
        for p in parts:
            if '@' in p:
                local, dom = p.split('@', 1)
                if local and dom and domain_re.match(dom):
                    ref_domain = '@' + dom
                    break

        normalized = []
        for p in parts:
            if '@' in p:
                local, dom = p.split('@', 1)
                local = local.strip()
                dom = dom.strip()
                if not local:
                    # skip invalid empty local
                    continue
                if dom and domain_re.match(dom):
                    # full valid email
                    normalized.append(f"{local}@{dom}")
                else:
                    # has '@' but missing/invalid domain: complete if we have a reference domain
                    if ref_domain:
                        normalized.append(local + ref_domain)
                    else:
                        # no reference domain—keep as-is (best effort)
                        normalized.append(f"{local}@{dom}" if dom else f"{local}@")
            else:
                # No '@': if looks like a local part, append reference domain if available
                if local_re.match(p):
                    if ref_domain:
                        normalized.append(p + ref_domain)
                    else:
                        normalized.append(p)
                else:
                    # Not a local—keep raw text
                    normalized.append(p)

        # De-dup while preserving order
        out = []
        seen = set()
        for e in normalized:
            if e not in seen and e:
                out.append(e)
                seen.add(e)

        return ';'.join(out).strip(' ;,')

    df_filtered['Email'] = df_filtered['Email'].apply(normalize_domains)

    # Merge with flags and metadata
    df_filtered = df_filtered.merge(final_df2, on='source_pdf', how='left')
    df_filtered = df_filtered.merge(
        final_df[['source_pdf', 'InvoiceDate', 'CustomerName', 'DOT']].drop_duplicates(),
        on='source_pdf', how='left'
    )

    # Deduplicate semicolon list again (defensive)
    def dedupe_semicolon_list(s):
        if not isinstance(s, str) or not s.strip():
            return s
        parts = [p.strip() for p in s.split(';') if p.strip()]
        seen = set()
        out = []
        for p in parts:
            if p not in seen:
                out.append(p)
                seen.add(p)
        return ';'.join(out)

    df_filtered['Email'] = df_filtered['Email'].apply(dedupe_semicolon_list)

    # ---------------------------------------------
    # Build "Email Subject Line" per Email group
    # Subject = "<CustomerName> - <DOT1,DOT2,...>" (if both exist)
    #         = "<CustomerName>" or "<DOTs>"       (fallbacks)
    # ---------------------------------------------
    def _build_subject_for_group(group: pd.DataFrame) -> str:
        cust = next((c for c in group['CustomerName'] if isinstance(c, str) and c.strip()), "")
        seen = set()
        dots = []
        for d in group['DOT']:
            if isinstance(d, str):
                d_clean = d.strip()
                if d_clean and d_clean not in seen:
                    seen.add(d_clean)
                    dots.append(d_clean)
        if cust and dots:
            return f"{cust} - {','.join(dots)}"
        elif cust:
            return cust
        elif dots:
            return ','.join(dots)
        else:
            return ""

    _subject_rows = []
    for email_val, grp in df_filtered.groupby('Email', sort=False):
        _subject_rows.append({
            'Email': email_val,
            'Email Subject Line': _build_subject_for_group(grp)
        })
    subject_map = pd.DataFrame(_subject_rows)
    df_filtered = df_filtered.merge(subject_map, on='Email', how='left')

    # ===========================
    # NEW: Short folder names -> "<number> - <first_email>"
    # ===========================
    df_filtered['Email'] = (
        df_filtered['Email']
        .fillna("")
        .astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
    )

    def _first_email(email_str: str) -> str:
        if not isinstance(email_str, str) or not email_str.strip():
            return ""
        parts = [p.strip() for p in email_str.split(';') if p.strip()]
        return parts[0] if parts else ""

    # Build group numbering in order of first appearance (stable within a run)
    seen = set()
    ordered_unique_emails = []
    for e in df_filtered['Email']:
        key = e.strip()
        if key and key not in seen:
            seen.add(key)
            ordered_unique_emails.append(key)

    email_to_number = {e: i + 1 for i, e in enumerate(ordered_unique_emails)}

    df_filtered['GroupNumber'] = df_filtered['Email'].map(email_to_number)
    df_filtered['FirstEmail'] = df_filtered['Email'].apply(_first_email)

    def _make_folder_name(num, first_email):
        # Empty -> no folder (goes to NoEmailFound)
        if pd.isna(num) or not str(first_email).strip():
            return ""
        # Optional cap (uncomment to enforce a shorter visible first_email)
        # first_email = first_email[:60]
        return safe_folder_name(f"{int(num)} - {first_email}")

    df_filtered['FinalFolderName'] = df_filtered.apply(
        lambda r: _make_folder_name(r['GroupNumber'], r['FirstEmail']), axis=1
    )

    # This column is optional; keep if you use it later
    if 'Email Verified? (Yes/No)' not in df_filtered.columns:
        df_filtered['Email Verified? (Yes/No)'] = ""

    # ===========================
    # Move PDFs into folders (instead of copying)
    # ===========================
    for _, row in df_filtered.iterrows():
        pdf_name = row['source_pdf']
        src_pdf_path = os.path.join(input_folder, pdf_name)

        dni_val = row.get('DoNotEmail', None)
        if isinstance(dni_val, str) and dni_val.strip():
            # Move to DoNotEmail
            move_with_collision(src_pdf_path, do_not_email_folder, pdf_name)
            continue

        composite = (row.get('FinalFolderName', '') or '').strip()
        if not composite:
            # Move to NoEmailFound
            move_with_collision(src_pdf_path, no_email_folder, pdf_name)
            continue

        folder_name = safe_folder_name(composite)
        email_folder = os.path.join(base_output_folder, folder_name)
        # move with collision handling
        move_with_collision(src_pdf_path, email_folder, pdf_name)

    # ===========================
    # Write Excel mapping
    # ===========================
    excel_path = os.path.join(base_output_folder, "Email_and_PDF_Mapping.xlsx")
    df_filtered.sort_values(by="FinalFolderName").to_excel(excel_path, index=False)

    print("✅ PDFs have been grouped (moved) by email into:", base_output_folder)
    print(f"🚫 DoNotEmail → {do_not_email_folder}")
    print(f"📂 NoEmailFound → {no_email_folder}")
    print(f"📒 Mapping created → {excel_path}")


# ===========================
# Program Wrapper
# ===========================
def ask_to_run_again() -> bool:
    try:
        while True:
            resp = input("\nRun again? (yes/no): ").strip().lower()
            if resp in {"y", "yes"}:
                return True
            if resp in {"n", "no"}:
                return False
            print("Please answer yes/no")
    except (EOFError, KeyboardInterrupt):
        print("\nNo input detected. Exiting.")
        return False

def wait_before_exit_prompt():
    try:
        input("Press Enter to exit...")
    except Exception:
        pass


if __name__ == "__main__":
    while True:
        try:
            main()
        except Exception as e:
            print("\n❌ An unexpected error occurred:")
            print(f"   {type(e).__name__}: {e}")
            print("\n--- Stack Trace ---")
            traceback.print_exc()
            print("-------------------\n")
            wait_before_exit_prompt()
            sys.exit(1)

        if not ask_to_run_again():
            print("👋 Done. Exiting.")
            sys.exit(0)