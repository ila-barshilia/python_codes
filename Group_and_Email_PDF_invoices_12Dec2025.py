
import sys
import traceback

def main():
    import os
    import pandas as pd
    import fitz
    from collections import defaultdict
    import shutil
    import re  # <<< added for targeted pattern handling
    import numpy as np

    # Configuration
    input_folder = input("Enter the path to where your PDF files are kept that you want to email to customers: ")
    base_output_folder = os.path.join(input_folder, "Grouped_PDFs")
    os.makedirs(base_output_folder, exist_ok=True)

    # >>> NEW: Special folder for PDFs with blank email
    no_email_folder = os.path.join(base_output_folder, "NoEmailFound")
    os.makedirs(no_email_folder, exist_ok=True)

    do_not_email_folder = os.path.join(base_output_folder, "DoNotEmail")
    os.makedirs(do_not_email_folder, exist_ok=True)

    # Store all PDFs' structured data
    all_dfs = []

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

    # ---------------- PROCESS ALL PDFs ----------------
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(input_folder, filename)

            df = extract_structured_text(pdf_path)
            df2 = group_by_dynamic_columns(df)

            df2["source_pdf"] = filename
            all_dfs.append(df2)

    final_df = pd.concat(all_dfs, ignore_index=True)

    def extract_right_concat(
        df: pd.DataFrame,
        keywords=("Job No", "P.O. #"),
        in_column: str | None = None,
        sep: str = " ",
        case_insensitive: bool = True
    ) -> pd.Series:

        cols = list(df.columns)
        second_last_idx = len(cols) - 2  # index of second-last column

        # Precompute lowercase keywords if needed
        kw_set = {k.lower() if case_insensitive else k for k in keywords}

        def _match(cell) -> bool:
            if cell is None or (isinstance(cell, float) and np.isnan(cell)):
                return False
            s = str(cell)
            s_cmp = s.lower() if case_insensitive else s
            return any(k in s_cmp for k in kw_set)

        def _concat_row(row: pd.Series) -> str | float:
            # Find keyword position
            if in_column is not None:
                if in_column not in df.columns:
                    raise ValueError(f"`in_column` '{in_column}' not in DataFrame columns.")
                pos = cols.index(in_column)
                if not _match(row[in_column]):
                    return np.nan
            else:
                # Search across the entire row for the first occurrence
                pos = None
                for j, c in enumerate(cols):
                    if _match(row[c]):
                        pos = j
                        break
                if pos is None:
                    return np.nan

            # Compute slice start (immediately to the right) and end (second-last col inclusive)
            start = pos + 1
            end = second_last_idx
            if start > end:
                return np.nan  # nothing to extract

            # Extract, drop None/NaN/empty, concatenate
            vals = row.iloc[start:end+1].tolist()
            cleaned = []
            for v in vals:
                if v is None or (isinstance(v, float) and np.isnan(v)):
                    continue
                s = str(v).strip()
                if s:  # exclude empty after trimming
                    cleaned.append(s)
            if not cleaned:
                return np.nan
            return sep.join(cleaned)

        return df.apply(_concat_row, axis=1)
    final_df["DoNotEmail"] = extract_right_concat(final_df, keywords=("Job No", "P.O. #"))

    final_df2 = final_df[['source_pdf', 'DoNotEmail']]
    final_df2 = final_df2[final_df2["DoNotEmail"].str.contains('Pay App|Adj Inv', case=False, na=False)]
    final_df2.drop_duplicates(subset=['source_pdf'], keep='first')

    # Work on a copy with a fresh RangeIndex so iloc works with your window logic
    fd = final_df.reset_index(drop=True)

    keyword = "Email"
    columns_to_check = [0, 1, 2]

    # Find rows where any of these columns contains 'Email' (case-insensitive)
    matches = fd.index[
        fd[columns_to_check]
        .astype(str)
        .apply(lambda row: any(keyword.lower() in str(val).lower() for val in row), axis=1)]

    # Include Email row + next 2 rows
    rows_to_include = []
    for pos in matches:  # pos are now positions 0..len(fd)-1
        end_pos = min(pos + 3, len(fd))
        rows_to_include.extend(range(pos, end_pos))

    rows_to_include = sorted(set(rows_to_include))
    df_filtered = fd.iloc[rows_to_include].reset_index(drop=True)

    # Remove unwanted rows
    df_filtered[0] = df_filtered[0].astype(str).str.replace("\u00A0", " ").str.strip()
    df_filtered = df_filtered[(df_filtered[0] != 'Customer Phone') & (df_filtered[0] != 'Customer Fax')]
    df_filtered.drop(columns="DoNotEmail", inplace=True)


    def combine_email_per_pdf_blocks(df_filtered: pd.DataFrame) -> pd.DataFrame:
        if 'source_pdf' not in df_filtered.columns:
            raise ValueError("Expected 'source_pdf' column in df_filtered.")
        src_idx = df_filtered.columns.get_loc('source_pdf')
        end_col = src_idx - 1  # last usable column index

        def norm(x):
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return ""
            return str(x).replace("\u00A0", " ").strip()

        search_cols = [0, 1, 2]
        markers = {"email", "none", "nan", ""}

        out_rows = []
        for pdf, grp in df_filtered.groupby('source_pdf', sort=False):
            email_pos = None
            first_match_row = None

            # Find first occurrence of 'Email'
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

    # Replace df_filtered for downstream steps
    df_filtered = df_email_by_pdf[['source_pdf', 'Email']]

    df_filtered = df_filtered.groupby('source_pdf')['Email'].agg(lambda x: ''.join(x)).reset_index()
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'\*+', '', regex=True).str.strip()

    # >>> NEW: Targeted move of only specific notes to Comment, and strip them from Email
    def extract_comments_and_clean(email_str):
        s = email_str or ""
        comments = []

        # Detect presence (case-insensitive) for standardized labels
        if re.search(r'\bUS\s*Mail\b', s, flags=re.IGNORECASE):
            comments.append("US Mail")
        if re.search(r'\bUSMail\b', s, flags=re.IGNORECASE):
            comments.append("USMail")
        if re.search(r'\bread\s+receipt\b', s, flags=re.IGNORECASE):
            comments.append("read receipt")
        if re.search(r'&\s*PM\b', s):
            comments.append("PM")

        # Remove only those occurrences (including common delimiters around them)
        # Slash before US Mail/USMail (e.g., '/US Mail', '/USMail')
        s = re.sub(r'(?i)\s*/\s*US\s*Mail', '', s)
        s = re.sub(r'(?i)\s*/\s*USMail', '', s)

        # 'read receipt' with or without parentheses
        s = re.sub(r'(?i)\s*\(\s*read\s+receipt\s*\)\s*', '', s)  # (read receipt)
        s = re.sub(r'(?i)\s*read\s+receipt\s*', '', s)            # read receipt

        # '& PM' or '&PM'
        s = re.sub(r'\s*&\s*PM\b', '', s)

        # Tidy spaces and commas after removals
        s = re.sub(r'\s+', ' ', s)
        s = re.sub(r'\s*,\s*', ', ', s)
        s = re.sub(r'(,\s*){2,}', ', ', s)
        s = s.strip(' ,')

        # De-duplicate comments preserving order
        seen = set()
        comments_unique = []
        for c in comments:
            if c not in seen:
                comments_unique.append(c)
                seen.add(c)

        return s, '; '.join(comments_unique)

    # Apply extraction
    clean_cols = df_filtered['Email'].apply(extract_comments_and_clean)
    df_filtered['Email'] = clean_cols.apply(lambda x: x[0])     # cleaned email string
    df_filtered['Comment'] = clean_cols.apply(lambda x: x[1])   # ONLY targeted notes
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'(?i)/\s*cc\s*([A-Za-z0-9._%+-]+)', r';\1', regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'(?i)/\s*cc\b', r';', regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'\s*/cc\s*', ';', flags=re.IGNORECASE, regex=True)
    df_filtered['Email'] = df_filtered['Email'].str.replace(r'/', ';', regex=True)


    def normalize_domains(email_str):
        if not isinstance(email_str, str) or not email_str.strip():
            return email_str

        # Unify commas to semicolons and trim
        s = re.sub(r'\s*,\s*', ';', email_str.strip())

        # Split on semicolons
        parts = [p.strip() for p in s.split(';') if p.strip()]
        if not parts:
            return email_str

        # Find domain from the first email-like token that contains '@'
        domain = None
        for p in parts:
            if '@' in p:
                at_idx = p.find('@')
                dom = p[at_idx:]  # includes '@'
                if len(dom) > 1:
                    domain = dom
                    break

        # If no domain found, do nothing
        if not domain:
            out = ';'.join(parts)
            out = re.sub(r'\s*;\s*', ';', out)
            out = re.sub(r';{2,}', ';', out)
            return out.strip(' ;,')

        normalized = []
        for i, p in enumerate(parts):
            if i == 0:
                # Keep the first token as-is (it carries the canonical domain)
                normalized.append(p)
                continue

            token = p.strip(' ,')

            if '@' in token:
                # Replace everything from '@' with the first domain
                local = token.split('@', 1)[0]
                normalized.append(local + domain)
            else:
                # Append domain if token is a valid local-part (no '@' present)
                # Accept letters/digits and . _ % + - ; no need to require a dot
                if re.match(r'^[A-Za-z0-9._%+-]+$', token):
                    normalized.append(token + domain)
                else:
                    normalized.append(token)  # leave non-email tokens

        out = ';'.join(normalized)
        out = re.sub(r'\s*;\s*', ';', out)
        out = re.sub(r';{2,}', ';', out)
        return out.strip(' ;,')


    # >>> NEW: Normalize domains across ';' separated tokens using the first email's domain
    df_filtered['Email'] = df_filtered['Email'].apply(normalize_domains)
    df_filtered = df_filtered.merge(final_df2, on='source_pdf', how='left')


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

    # Ensure folders exist (if not already created above)
    os.makedirs(base_output_folder, exist_ok=True)
    os.makedirs(no_email_folder, exist_ok=True)
    os.makedirs(do_not_email_folder, exist_ok=True)  # <<< make sure this exists

    # ---------------- GROUP PDFs BY EMAIL (with DoNotEmail precedence) ----------------
    for _, row in df_filtered.iterrows():
        pdf_name = row['source_pdf']
        src_pdf_path = os.path.join(input_folder, pdf_name)

        # 1) >>> PRIORITY RULE: If DoNotEmail is present (non-NaN/non-empty), put PDF in DoNotEmail folder and skip further logic
        dni_val = row.get('DoNotEmail', None)
        dni_has_value = isinstance(dni_val, str) and dni_val.strip() != ""
        if (dni_val is not None) and (not isinstance(dni_val, float)) and dni_has_value:
            dest_pdf_path = os.path.join(do_not_email_folder, pdf_name)
            if not os.path.exists(dest_pdf_path):  # Avoid duplicate copies
                shutil.copy(src_pdf_path, dest_pdf_path)
            # Skip to next PDF; DoNotEmail takes precedence over NoEmailFound and email grouping
            continue

        # 2) Existing behavior: split Email tokens and determine grouping
        # Normalize split on semicolon OR comma and drop empties
        email_raw = row.get('Email', "")
        emails = [e.strip() for e in re.split(r'[;,]', str(email_raw)) if e.strip()]

        # If NO '@' appears anywhere in the cleaned Email string, move to NoEmailFound (keep original filename)
        emails_with_at = [e for e in emails if '@' in e]
        if not emails_with_at:
            dest_pdf_path = os.path.join(no_email_folder, pdf_name)
            if not os.path.exists(dest_pdf_path):  # Avoid duplicate copies
                shutil.copy(src_pdf_path, dest_pdf_path)
            continue

        # Use cleaned emails (with '@') for folder grouping
        for email in emails_with_at:
            email_folder = os.path.join(base_output_folder, email)
            os.makedirs(email_folder, exist_ok=True)

            dest_pdf_path = os.path.join(email_folder, pdf_name)
            if not os.path.exists(dest_pdf_path):  # Avoid duplicate copies
                shutil.copy(src_pdf_path, dest_pdf_path)

    # Write mapping
    excel_path = os.path.join(base_output_folder, "Email_and_PDF_Mapping.xlsx")
    df_filtered_sorted = df_filtered.sort_values(by="Email")
    df_filtered_sorted.to_excel(excel_path, index=False)

    print("✅ PDFs have been grouped by email into folders inside:", base_output_folder)
    print(f"🚫 PDFs flagged with DoNotEmail were placed in: {do_not_email_folder}")
    print(f"📂 PDFs with blank Email were placed in: {no_email_folder}")



def ask_to_run_again() -> bool:
    try:
        while True:
            resp = input("\nDo you have more grouping tasks to run for other split PDF invoices? (yes/no): ").strip().lower()
            if resp in {"y", "yes"}:
                return True
            elif resp in {"n", "no"}:
                return False
            else:
                print("Please answer 'yes' or 'no'.")
    except (EOFError, KeyboardInterrupt):
        # If input is not available (e.g., piped execution), default to terminating
        print("\nNo input detected. Exiting.")
        return False



def wait_before_exit_prompt():
    try:
        input("⏸ Press Enter to exit...")
    except (EOFError, KeyboardInterrupt):
        # Non-interactive environment or user interrupted; just proceed to exit
        pass

if __name__ == "__main__":
    # Global loop + error handling
    while True:
        try:
            main()
        except Exception as e:
            # Print a clear error message and full stack trace before closing
            print("\n❌ An unexpected error occurred:")
            print(f"   {type(e).__name__}: {e}")
            print("\n--- Stack Trace ---")
            traceback.print_exc()
            print("-------------------\n")

            # >>> Pause so you can capture the error
            wait_before_exit_prompt()

            # Program closes after showing error (per your request)
            sys.exit(1)

        # Ask the user if they want to run again
        if not ask_to_run_again():
            print("👋 Done. Exiting.")
            sys.exit(0)
        # Otherwise, loop continues and `main()` runs anew

