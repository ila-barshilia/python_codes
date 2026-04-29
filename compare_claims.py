 
import argparse
import re
from decimal import Decimal, InvalidOperation
import pandas as pd

# ----------------------------
# Helpers for normalization
# ----------------------------

CURRENCY_STRIP_RE = re.compile(r"[^\d\-\.\(\),]")


def to_decimal_exact(x):
    """
    Convert a value to Decimal for exact comparison of financial fields.
    - Supports currency-like strings (₹, $, commas) and parentheses negatives.
    - Returns pd.NA if not parseable.
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
# Core comparison logic
# ----------------------------

def prepare_frame(df, acme_key, carrier_key, cols_to_compare, financial_cols, case_sensitive):
    # Validate
    required = [acme_key, carrier_key] + cols_to_compare
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in input: {missing}")

    out = df[required].copy()

    # Normalize financials
    for c in cols_to_compare:
        if c in financial_cols:
            out[c] = out[c].apply(to_decimal_exact)
        else:
            # text normalization for non-financials
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
    Attempt to pair remaining (unpaired) rows by the provided key + occurrence index.
    Returns: pairs_df, new_used_left_idx, new_used_right_idx
    """
    l_remaining = left.loc[~left.index.isin(used_left_idx)].copy()
    r_remaining = right.loc[~right.index.isin(used_right_idx)].copy()

    # Preserve original indices
    l_remaining = l_remaining.reset_index().rename(columns={"index": "_idx_A"})
    r_remaining = r_remaining.reset_index().rename(columns={"index": "_idx_B"})

    merged = l_remaining.merge(
        r_remaining,
        on=[key_col, occ_col],
        how="inner",
        suffixes=("_A", "_B"),
        indicator=False,
    )

    if merged.empty:
        return merged, used_left_idx, used_right_idx

    merged["Pairing_Method"] = label

    # Update used indices sets
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
    """
    Build a long-form table of differences for the paired rows only.
    paired_df contains columns: ... <col>_A, <col>_B, Pairing_Method, match keys and occ columns.
    """
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


def run_compare(args):
    # Column names (defaults set to Ila's)
    acme_key = args.acme_key
    carrier_key = args.carrier_key
    cols_to_compare = [c.strip() for c in args.cols.split(",") if c.strip()]

    # Validate 5 columns expected (4 financial + 1 text)
    if len(cols_to_compare) != 5:
        raise ValueError(f"--cols must specify exactly 5 columns to compare; got {len(cols_to_compare)}")

    financial_cols = set([c.strip() for c in args.financial_cols.split(",") if c.strip()])
    missing_fin = [c for c in financial_cols if c not in cols_to_compare]
    if missing_fin:
        raise ValueError(f"These financial columns are not in --cols: {missing_fin}")

    # Read files
    df1 = pd.read_excel(args.file1, sheet_name=args.sheet1, engine="openpyxl")
    df2 = pd.read_excel(args.file2, sheet_name=args.sheet2, engine="openpyxl")

    # Prepare frames with normalization and occurrence indices
    left = prepare_frame(df1, acme_key, carrier_key, cols_to_compare, financial_cols, not args.case_insensitive)
    right = prepare_frame(df2, acme_key, carrier_key, cols_to_compare, financial_cols, not args.case_insensitive)

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
    if pairs_a is None and pairs_b is None:
        pairs = pd.DataFrame()
    else:
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
    p = argparse.ArgumentParser(description="Compare two Excel files per claim using Acme then Carrier keys; 4 financial + 1 text column comparison with exact financial checks.")
    p.add_argument("--file1", required=True, help="Path to first Excel file (File A).")
    p.add_argument("--file2", required=True, help="Path to second Excel file (File B).")
    p.add_argument("--sheet1", default=0, help="Sheet name or index for File A (default: first sheet).")
    p.add_argument("--sheet2", default=0, help="Sheet name or index for File B (default: first sheet).")
    p.add_argument("--acme-key", dest="acme_key", default="Acme Claim #", help="Column name for Acme Claim # (primary key).")
    p.add_argument("--carrier-key", dest="carrier_key", default="Carrier Claim #", help="Column name for Carrier Claim # (secondary key).")
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
    p.add_argument("--case-insensitive", dest="case_insensitive", action="store_true", help="Use case-insensitive text comparison for non-financial fields (e.g., Status: O/C).")
    p.add_argument("--case-sensitive", dest="case_insensitive", action="store_false", help="Use case-sensitive text comparison (default).")
    p.set_defaults(case_insensitive=False)
    return p.parse_args()


if __name__ == '__main__':
    args = parse_args()
    run_compare(args)
