import argparse
import re
from decimal import Decimal, InvalidOperation
import pandas as pd
import sys

# ----------------------------
# Helpers: numeric / text normalization
# ----------------------------

CURRENCY_STRIP_RE = re.compile(r"[^\d\-\.\(\),]")

def to_decimal_exact(x):
    """
    Convert a value to Decimal for exact comparison of financial fields.
    Supports currency symbols, commas, and (parentheses) negatives.
    """
    if pd.isna(x):
        return pd.NA
    try:
        if isinstance(x, (int, float)):
            return Decimal(str(x))
        s = str(x).strip()
        if s == "":
            return pd.NA
        s = CURRENCY_STRIP_RE.sub("", s)
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1]
        s = s.replace(",", "")
        d = Decimal(s)
        return -d if neg else d
    except (InvalidOperation, ValueError):
        return pd.NA

def normalize_text(x, case_sensitive=True):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    return s if case_sensitive else s.lower()

def add_occurrence_index(df, key_col, occ_name):
    df = df.copy()
    df[occ_name] = df.groupby(key_col, dropna=False).cumcount()
    return df

# ----------------------------
# Sheet name resolution (whitespace-insensitive)
# ----------------------------

def _collapse_internal_spaces(name: str) -> str:
    return " ".join(str(name).split())

def _remove_all_spaces(name: str) -> str:
    return str(name).replace(" ", "")

def resolve_sheet_name(xls: pd.ExcelFile, requested):
    """
    Resolve a sheet name by trying:
    - exact / strip / collapsed spaces / no spaces
    - case-insensitive variants of the above
    """
    sheets = list(xls.sheet_names)
    req = str(requested)
    candidates = [req, req.strip(), _collapse_internal_spaces(req), _remove_all_spaces(req)]

    for cand in candidates:
        if cand in sheets:
            return cand

    lower_map = {s.lower(): s for s in sheets}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]

    def key_space_insensitive(s: str) -> str:
        return _remove_all_spaces(_collapse_internal_spaces(s)).lower()

    norm_map = {key_space_insensitive(s): s for s in sheets}
    key = key_space_insensitive(req)
    if key in norm_map:
        return norm_map[key]

    raise ValueError(f"Sheet '{requested}' not found. Available sheets: {sheets}")

# ----------------------------
# Column name normalization & aliasing
# ----------------------------

def canon(name: str) -> str:
    """
    Canonicalize a column header to compare loosely:
    - lowercase
    - strip
    - collapse spaces
    - remove punctuation and special symbols (#, :, /, etc.)
    """
    if name is None:
        return ""
    s = str(name).strip().lower()
    s = " ".join(s.split())  # collapse internal spaces
    # remove punctuation-like characters
    s = re.sub(r"[^\w\s]", "", s)  # keep letters/numbers/underscore/space
    s = s.replace("_", " ")
    s = " ".join(s.split())
    return s

def build_alias_map(actual_cols, expected_cols, debug=False):
    """
    Create a mapping from expected -> actual by canonical form.
    If multiple actuals map to the same canon, prefer exact/closest matches.
    """
    # Index actual columns by their canonical form
    actual_by_canon = {}
    for ac in actual_cols:
        ac_can = canon(ac)
        actual_by_canon.setdefault(ac_can, []).append(ac)

    mapping = {}
    missing = []

    for exp in expected_cols:
        exp_can = canon(exp)
        candidates = actual_by_canon.get(exp_can)
        if candidates:
            # If multiple candidates share canon, choose the one with minimal edit distance by length diff
            # (light heuristic; usually single candidate)
            chosen = sorted(candidates, key=lambda c: abs(len(str(c)) - len(str(exp))))[0]
            mapping[exp] = chosen
        else:
            # Try looser match: remove extra spaces-only canon (already done), nothing found -> missing
            missing.append(exp)

    if debug:
        print(">> DEBUG: Column alias map (expected -> actual):", file=sys.stderr)
        for k, v in mapping.items():
            print(f"   {k!r} -> {v!r}", file=sys.stderr)
        if missing:
            print(">> DEBUG: Still missing expected columns:", missing, file=sys.stderr)

    return mapping, missing

def rename_with_aliases(df, alias_map):
    """
    Returns a copy of df with columns renamed per alias_map (expected->actual reversed).
    We need to convert from actual->expected for rename().
    """
    if not alias_map:
        return df
    # Build reverse: actual -> expected
    actual_to_expected = {actual: expected for expected, actual in alias_map.items()}
    return df.rename(columns=actual_to_expected)

# ----------------------------
# Core comparison logic
# ----------------------------

def prepare_frame(df, acme_key, carrier_key, cols_to_compare, financial_cols, case_sensitive):
    # Validate after aliasing/renaming has been applied in caller
    required = [acme_key, carrier_key] + cols_to_compare
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in input: {missing}")

    out = df[required].copy()

    # Normalize financials and text
    for c in cols_to_compare:
        if c in financial_cols:
            out[c] = out[c].apply(to_decimal_exact)
        else:
            if pd.api.types.is_object_dtype(out[c].dtype):
                out[c] = out[c].apply(lambda v: normalize_text(v, case_sensitive))
            else:
                out[c] = out[c].apply(lambda v: normalize_text(v, case_sensitive) if isinstance(v, str) else v)

    # Add occurrence indices for each key
    out = add_occurrence_index(out, acme_key, "_occ_acme")
    out = add_occurrence_index(out, carrier_key, "_occ_carrier")
    return out

def pair_by_key(left, right, key_col, occ_col, used_left_idx, used_right_idx, label):
    """
    Pair remaining rows by (key + occurrence index)
    """
    l_remaining = left.loc[~left.index.isin(used_left_idx)].copy()
    r_remaining = right.loc[~right.index.isin(used_right_idx)].copy()

    # Preserve original indices
    l_remaining = l_remaining.reset_index().rename(columns={"index": "_idx_A"})
    r_remaining = r_remaining.reset_index().rename(columns={"index": "_idx_B"})

    merged = l_remaining.merge(
        r_remaining, on=[key_col, occ_col], how="inner", suffixes=("_A", "_B"), indicator=False
    )
    if merged.empty:
        return merged, used_left_idx, used_right_idx

    merged["Pairing_Method"] = label
    used_left_idx.update(merged["_idx_A"].tolist())
    used_right_idx.update(merged["_idx_B"].tolist())
    return merged, used_left_idx, used_right_idx

def compare_rows(row, cols_to_compare):
    for c in cols_to_compare:
        va = row.get(f"{c}_A")
        vb = row.get(f"{c}_B")
        if (pd.isna(va) and pd.isna(vb)):
            continue
        if va != vb:
            return False
    return True

def build_differences_long(paired_df, cols_to_compare, match_key_label, occ_label):
    records = []
    for _, r in paired_df.iterrows():
        for c in cols_to_compare:
            va = r.get(f"{c}_A")
            vb = r.get(f"{c}_B")
            if not ((pd.isna(va) and pd.isna(vb)) or (va == vb)):
                records.append({
                    "Pairing_Method": r["Pairing_Method"],
                    match_key_label: r[match_key_label],
                    occ_label: r[occ_label],
                    "Column": c,
                    "Value_in_File1": va,
                    "Value_in_File2": vb,
                })
    if not records:
        return pd.DataFrame(columns=["Pairing_Method", match_key_label, occ_label, "Column", "Value_in_File1", "Value_in_File2"])
    df_long = pd.DataFrame(records)
    return df_long.sort_values(["Pairing_Method", match_key_label, occ_label, "Column"], kind="stable")

def read_sheet_resolved(file_path, sheet_requested, header_row_index, debug=False):
    """
    Open Excel, resolve sheet name robustly, read with a given header row index (0-based).
    """
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    actual = resolve_sheet_name(xls, sheet_requested) if isinstance(sheet_requested, str) else sheet_requested
    if debug:
        print(f">> DEBUG: File '{file_path}' resolved sheet -> '{actual}'", file=sys.stderr)
    df = pd.read_excel(xls, sheet_name=actual, engine='openpyxl', header=header_row_index)
    if debug:
        print(f">> DEBUG: Columns in '{file_path}' / '{actual}': {list(df.columns)}", file=sys.stderr)
        print(f">> DEBUG: Head(3):\n{df.head(3)}", file=sys.stderr)
    return df

def run_compare(args):
    # Column names (defaults set to Ila's)
    acme_key = args.acme_key
    carrier_key = args.carrier_key
    cols_to_compare = [c.strip() for c in args.cols.split(",") if c.strip()]
    debug = args.debug

    # Validate 5 columns expected (4 financial + 1 text)
    if len(cols_to_compare) != 5:
        raise ValueError(f"--cols must specify exactly 5 columns to compare; got {len(cols_to_compare)}")

    financial_cols = set([c.strip() for c in args.financial_cols.split(",") if c.strip()])
    missing_fin = [c for c in financial_cols if c not in cols_to_compare]
    if missing_fin:
        raise ValueError(f"These financial columns are not in --cols: {missing_fin}")

    # Convert 1-based header row to 0-based index for pandas
    header_row_index = max(args.header_row - 1, 0)

    # Read files with sheet name resolution + header row
    df1 = read_sheet_resolved(args.file1, args.sheet1, header_row_index, debug=debug)
    df2 = read_sheet_resolved(args.file2, args.sheet2, header_row_index, debug=debug)

    # Build alias maps so minor header differences don't break matching
    expected_all = [acme_key, carrier_key] + cols_to_compare

    alias_map_1, missing_1 = build_alias_map(df1.columns, expected_all, debug=debug)
    alias_map_2, missing_2 = build_alias_map(df2.columns, expected_all, debug=debug)

    df1_named = rename_with_aliases(df1, alias_map_1)
    df2_named = rename_with_aliases(df2, alias_map_2)

    # If still missing after aliasing, raise with diagnostics
    still_missing_1 = [c for c in expected_all if c not in df1_named.columns]
    still_missing_2 = [c for c in expected_all if c not in df2_named.columns]

    if still_missing_1 or still_missing_2:
        msg = []
        if still_missing_1:
            msg.append(f"File1 missing after normalization: {still_missing_1}\nColumns seen: {list(df1.columns)}")
        if still_missing_2:
            msg.append(f"File2 missing after normalization: {still_missing_2}\nColumns seen: {list(df2.columns)}")
        raise ValueError("\n".join(msg))

    # Prepare frames with normalization and occurrence indices
    left = prepare_frame(df1_named, acme_key, carrier_key, cols_to_compare, financial_cols, not args.case_insensitive)
    right = prepare_frame(df2_named, acme_key, carrier_key, cols_to_compare, financial_cols, not args.case_insensitive)

    # Stage A: pair by Acme Claim #
    used_left_idx, used_right_idx = set(), set()
    acme_occ = "_occ_acme"
    carrier_occ = "_occ_carrier"

    pairs_a, used_left_idx, used_right_idx = pair_by_key(
        left, right, acme_key, acme_occ, used_left_idx, used_right_idx, label=acme_key
    )

    # Stage B: pair remaining by Carrier Claim #
    pairs_b, used_left_idx, used_right_idx = pair_by_key(
        left, right, carrier_key, carrier_occ, used_left_idx, used_right_idx, label=carrier_key
    )

    # Concatenate all pairs
    parts = []
    if pairs_a is not None and not pairs_a.empty:
        parts.append(pairs_a)
    if pairs_b is not None and not pairs_b.empty:
        parts.append(pairs_b)
    pairs = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()

    # If no pairs found, just output unmatched
    if pairs.empty:
        only_in_A = left.loc[~left.index.isin(used_left_idx)].copy()
        only_in_B = right.loc[~right.index.isin(used_right_idx)].copy()
        with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
            only_in_A.to_excel(writer, sheet_name="Only_in_File1", index=False)
            only_in_B.to_excel(writer, sheet_name="Only_in_File2", index=False)
        print("No pairs found. Unmatched rows exported.")
        return

    # Determine which key/occ was used for each pair for reporting
    def select_match_key(row):
        if row["Pairing_Method"] == acme_key:
            return row[f"{acme_key}_A"], row[f"{acme_occ}_A"], acme_key, acme_occ
        else:
            return row[f"{carrier_key}_A"], row[f"{carrier_occ}_A"], carrier_key, carrier_occ

    pairs["_match_key_value"], pairs["_match_occ"], _, _ = zip(*pairs.apply(select_match_key, axis=1))

    # Compute exact matches vs differences
    pairs["is_exact_match"] = pairs.apply(lambda r: compare_rows(r, cols_to_compare), axis=1)

    # Build exact sheet (show A side since identical)
    exact_cols = [
        "Pairing_Method",
        "_match_key_value",
        "_match_occ",
        acme_key + "_A",
        carrier_key + "_A",
    ] + [f"{c}_A" for c in cols_to_compare]

    exact_out = pairs[pairs["is_exact_match"]][exact_cols].copy()
    rename_map = {
        "_match_key_value": "Matched_Key_Value",
        "_match_occ": "Occurrence_Index",
        acme_key + "_A": acme_key,
        carrier_key + "_A": carrier_key,
    }
    for c in cols_to_compare:
        rename_map[f"{c}_A"] = c
    exact_out.rename(columns=rename_map, inplace=True)

    # Differences long-form
    diffs_long = build_differences_long(
        pairs[~pairs["is_exact_match"]],
        cols_to_compare,
        match_key_label="_match_key_value",
        occ_label="_match_occ",
    )
    diffs_long.rename(columns={"_match_key_value": "Matched_Key_Value", "_match_occ": "Occurrence_Index"}, inplace=True)

    # Unmatched after both pairing stages
    only_in_A = left.loc[~left.index.isin(used_left_idx)].copy()
    only_in_B = right.loc[~right.index.isin(used_right_idx)].copy()

    unmatched_cols = [acme_key, carrier_key, "_occ_acme", "_occ_carrier"] + cols_to_compare
    only_in_A = only_in_A[unmatched_cols]
    only_in_B = only_in_B[unmatched_cols]

    # Save report
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        exact_out.to_excel(writer, sheet_name="Exact_Matches", index=False)
        diffs_long.to_excel(writer, sheet_name="Differences_Detail", index=False)
        only_in_A.to_excel(writer, sheet_name="Only_in_File1", index=False)
        only_in_B.to_excel(writer, sheet_name="Only_in_File2", index=False)

    print("✅ Comparison complete.")
    print(f"  Exact matches: {len(exact_out):>6}")
    print(f"  Diff pairs   : {len(pairs) - len(exact_out):>6}  (see Differences_Detail)")
    print(f"  Only in A    : {len(only_in_A):>6}")
    print(f"  Only in B    : {len(only_in_B):>6}")
    print(f"  Report saved to: {args.output}")

def parse_args():
    p = argparse.ArgumentParser(
        description=(
            "Compare two Excel files per claim (Acme then Carrier key). "
            "4 financial + 1 text column comparison with exact financial checks. "
            "Sheet names are whitespace-insensitive. Supports header row selection and debug."
        )
    )
    p.add_argument("--file1", required=True, help="Path to first Excel file (File A).")
    p.add_argument("--file2", required=True, help="Path to second Excel file (File B).")
    p.add_argument("--sheet1", default=0, help="Sheet name or index for File A (e.g., 'Acme Data').")
    p.add_argument("--sheet2", default=0, help="Sheet name or index for File B (e.g., 'CarrierData').")
    p.add_argument("--header-row", type=int, default=1,
                   help="1-based row number containing headers in each sheet (default: 1 = first row).")
    p.add_argument("--acme-key", dest="acme_key", default="Acme Claim #",
                   help="Column name for Acme Claim # (primary key).")
    p.add_argument("--carrier-key", dest="carrier_key", default="Carrier Claim #",
                   help="Column name for Carrier Claim # (secondary key).")
    p.add_argument(
        "--cols",
        default="Carrier Net Incurred,Carrier Paid,Carrier Reserve,Recovered,Status: O/C",
        help="Comma-separated list of exactly 5 columns to compare (4 financial + 1 text).",
    )
    p.add_argument(
        "--financial-cols",
        default="Carrier Net Incurred,Carrier Paid,Carrier Reserve,Recovered",
        help="Comma-separated subset of --cols that are financial (exact Decimal match).",
    )
    p.add_argument("--output", default="claims_comparison_report.xlsx", help="Path for output Excel report.")
    p.add_argument("--case-insensitive", dest="case_insensitive", action="store_true",
                   help="Use case-insensitive text comparison for non-financial fields (e.g., Status: O/C).")
    p.add_argument("--case-sensitive", dest="case_insensitive", action="store_false",
                   help="Use case-sensitive text comparison (default).")
    p.add_argument("--debug", action="store_true", help="Print resolved sheet names, columns, and sample rows.")
    p.set_defaults(case_insensitive=False)
    return p.parse_args()

if __name__ == '__main__':
    args = parse_args()
    run_compare(args)