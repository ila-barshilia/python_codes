#!/usr/bin/env python3
"""
Send grouped PDFs via email based on a mapping Excel.

Workflow
--------
1) Ask user for the *same base input folder* that contains the PDFs and the `Grouped_PDFs` subfolder.
2) Load `Grouped_PDFs/Email_and_PDF_Mapping.xlsx`.
3) Filter rows where:
   - `Email Verified? (Yes/No)` == 'Yes'
   - and `FinalFolderName` (or fallback to `CompositeEmailFolder`) is not blank.
4) For each such row, gather all PDFs inside that folder and compose an email.
5) Default mode: create an .eml draft per folder (no credentials required).
   Optional: --send to actually send via SMTP using OAuth2 or app password.

Notes
-----
- Draft mode produces portable .eml files in `Grouped_PDFs/OutgoingDrafts` you can open/send via Outlook.
- SMTP send mode requires valid credentials and modern auth (OAuth2). Basic auth is deprecated for M365.
- Subject/Body can be supplied per-row via columns `Subject` and `Body`. If absent, defaults are used.
- Supports multiple recipients separated by ';' or ','.

"""

import os
import sys
import re
import smtplib
import ssl
from pathlib import Path
from typing import List, Optional
import pandas as pd
from email.message import EmailMessage
from datetime import datetime

# --- Constants ---
MAPPING_FILENAME = "Email_and_PDF_Mapping.xlsx"
GROUPED_FOLDER_NAME = "Grouped_PDFs"
OUTGOING_DRAFTS_DIR = "OutgoingDrafts"
LOG_FILENAME = "SendLog.csv"

# Default templates (used if mapping doesn't have Subject/Body columns)
DEFAULT_SUBJECT = "Your documents from {folder_name} ({pdf_count} attachment(s))"
DEFAULT_BODY = (
    "Hi,\n\n"
    "Please find attached the documents for your records.\n"
    "Folder: {folder_name}\n"
    "Attachments: {pdf_count}\n\n"
    "Regards,\n"
    "{sender_name}"
)

# Simple email validation
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")


def is_valid_email(s: str) -> bool:
    return bool(s and EMAIL_RE.match(s.strip()))


def split_emails(s: str) -> List[str]:
    if not isinstance(s, str):
        return []
    tokens = [t.strip() for t in re.split(r"[;,]", s) if t.strip()]
    # Dedupe while preserving order and keep only valid
    seen = set()
    out = []
    for t in tokens:
        # salvage email substring if needed
        m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", t)
        if m:
            e = m.group(0)
            if e not in seen:
                out.append(e)
                seen.add(e)
    return out


def safe_path_component(name: str) -> str:
    if not isinstance(name, str):
        name = str(name or "")
    # Replace illegal path chars but keep '@', '.', ';'
    name = re.sub(r'[<>:"/\\|?*]', '_', name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name[:240] if len(name) > 240 else name


def build_message(sender: str, recipients: List[str], subject: str, body: str,
                  attachments: List[Path]) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    for ap in attachments:
        # Try to guess subtype
        subtype = "pdf" if ap.suffix.lower() == ".pdf" else None
        with open(ap, "rb") as f:
            data = f.read()
        if subtype == "pdf":
            msg.add_attachment(data, maintype="application", subtype="pdf", filename=ap.name)
        else:
            # generic binary
            msg.add_attachment(data, maintype="application", subtype="octet-stream", filename=ap.name)
    return msg


def save_eml(msg: EmailMessage, dest_dir: Path, base_name: str) -> Path:
    dest_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fn = safe_path_component(f"{base_name}_{ts}.eml")
    eml_path = dest_dir / fn
    with open(eml_path, "wb") as f:
        f.write(msg.as_bytes())
    return eml_path


def send_smtp(msg: EmailMessage, smtp_server: str, smtp_port: int,
              username: str, password: Optional[str] = None, use_tls: bool = True) -> None:
    """
    Basic SMTP send. NOTE: For Microsoft 365, basic auth is deprecated.
    Prefer OAuth2 or app passwords. This function is provided for environments
    where SMTP AUTH is enabled.
    """
    if use_tls:
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls(context=context)
            server.login(username, password or "")
            server.send_message(msg)
    else:
        with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
            server.login(username, password or "")
            server.send_message(msg)


def main():
    base_input = input("Enter the SAME base input folder used for grouping (contains Grouped_PDFs): ").strip()
    if not base_input:
        print("No folder provided. Exiting.")
        sys.exit(1)

    base_path = Path(base_input)
    grouped_path = base_path / GROUPED_FOLDER_NAME
    mapping_path = grouped_path / MAPPING_FILENAME

    if not grouped_path.exists():
        print(f"❌ Missing '{GROUPED_FOLDER_NAME}' under: {base_path}")
        sys.exit(1)
    if not mapping_path.exists():
        print(f"❌ Mapping file not found: {mapping_path}")
        sys.exit(1)

    # Load mapping
    try:
        df = pd.read_excel(mapping_path, engine="openpyxl")
    except Exception as e:
        print(f"❌ Failed to read mapping Excel: {e}")
        sys.exit(1)

    # Normalize column names
    cols_norm = {c: c.strip() for c in df.columns}
    df.rename(columns=cols_norm, inplace=True)

    # Required columns
    required_cols = ["source_pdf", "Email", "Email Verified? (Yes/No)"]
    for rc in required_cols:
        if rc not in df.columns:
            print(f"❌ Required column missing in mapping: {rc}")
            sys.exit(1)

    # Folder name column (prefer FinalFolderName, else CompositeEmailFolder)
    folder_col = None
    if "FinalFolderName" in df.columns:
        folder_col = "FinalFolderName"
    else:
        print("❌ Neither 'FinalFolderName' nor 'CompositeEmailFolder' present in mapping.")
        sys.exit(1)

    # Optional per-row subject/body
    has_subject = "Subject" in df.columns
    has_body = "Body" in df.columns

    # Mode: drafts or send
    mode = input("Mode: type 'send' to email via SMTP, or press Enter for DRAFTS (.eml files): ").strip().lower()
    do_send = (mode == "send")

    # If sending, collect SMTP settings
    smtp_server = smtp_port = smtp_user = smtp_pass = None
    if do_send:
        smtp_server = input("SMTP server (e.g., smtp.office365.com): ").strip() or "smtp.office365.com"
        smtp_port_str = input("SMTP port (e.g., 587): ").strip() or "587"
        try:
            smtp_port = int(smtp_port_str)
        except ValueError:
            print("❌ Invalid port. Use a number like 587.")
            sys.exit(1)
        smtp_user = input("SMTP username (your email address): ").strip()
        smtp_pass = input("SMTP password/app password (leave blank if using OAuth in your environment): ").strip()
        if not smtp_user:
            print("❌ SMTP username required to send.")
            sys.exit(1)

    # Sender name and address
    sender_email = input("From email address (e.g., you@company.com): ").strip()
    sender_name = input("Sender display name (optional): ").strip() or "Sender"

    # Results log
    log_rows = []

    
    # Filter rows: Verified == 'Yes', folder present, and DoNotEmail blank
    verified_yes = df["Email Verified? (Yes/No)"].astype(str).str.strip().str.lower() == "yes"
    folder_present = df[folder_col].astype(str).str.strip() != ""
    # Treat NaN or empty strings as "blank"
    do_not_email_blank = df["DoNotEmail"].isna() | (df["DoNotEmail"].astype(str).str.strip() == "")

    filt = verified_yes & folder_present & do_not_email_blank
    df_send = df.loc[filt].copy()


    if df_send.empty:
        print("ℹ️ No rows qualified for sending (Verified=Yes and folder name present).")

    for idx, row in df_send.iterrows():
        folder_name_raw = str(row[folder_col]).strip()
        email_raw = str(row.get("Email", "")).strip()
        recipients = split_emails(email_raw)

        if not recipients:
            log_rows.append({"Folder": folder_name_raw, "Status": "Skipped - No valid recipients", "Detail": email_raw})
            continue

        # Build folder path and collect PDFs
        email_folder_path = grouped_path / safe_path_component(folder_name_raw)
        if not email_folder_path.exists():
            log_rows.append({"Folder": folder_name_raw, "Status": "Skipped - Folder not found", "Detail": str(email_folder_path)})
            continue

        pdfs = sorted([p for p in email_folder_path.glob("*.pdf") if p.is_file()])
        if not pdfs:
            log_rows.append({"Folder": folder_name_raw, "Status": "Skipped - No PDFs in folder", "Detail": str(email_folder_path)})
            continue

        # Subject/body per-row or defaults
        subject_tpl = str(row["Subject"]).strip() if has_subject and pd.notna(row["Subject"]) else DEFAULT_SUBJECT
        body_tpl = str(row["Body"]).strip() if has_body and pd.notna(row["Body"]) else DEFAULT_BODY

        subject = subject_tpl.format(folder_name=folder_name_raw, pdf_count=len(pdfs))
        body = body_tpl.format(folder_name=folder_name_raw, pdf_count=len(pdfs), sender_name=sender_name)

        # Build message
        msg = build_message(sender=sender_email, recipients=recipients, subject=subject, body=body, attachments=pdfs)

        try:
            if do_send:
                send_smtp(msg, smtp_server=smtp_server, smtp_port=smtp_port, username=smtp_user, password=smtp_pass)
                log_rows.append({"Folder": folder_name_raw, "Status": "SENT", "Detail": ", ".join(recipients)})
            else:
                drafts_dir = grouped_path / OUTGOING_DRAFTS_DIR / safe_path_component(folder_name_raw)
                eml_path = save_eml(msg, drafts_dir, base_name="draft")
                log_rows.append({"Folder": folder_name_raw, "Status": "DRAFT_CREATED", "Detail": str(eml_path)})
        except Exception as e:
            log_rows.append({"Folder": folder_name_raw, "Status": "ERROR", "Detail": str(e)})

    # Write log
    log_df = pd.DataFrame(log_rows)
    log_path = grouped_path / LOG_FILENAME
    try:
        log_df.to_csv(log_path, index=False)
        print(f"\n📒 Log written: {log_path}")
    except Exception as e:
        print(f"❌ Failed to write log: {e}")

    if not df_send.empty:
        print(f"\n✅ Completed processing {len(df_send)} row(s).")
        if not do_send:
            print(f"Draft emails (.eml) are in: {grouped_path / OUTGOING_DRAFTS_DIR}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted.")
