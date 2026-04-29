import os
import re
import shutil
import tempfile
import pandas as pd
import win32com.client as win32

# =========================
# Configuration
# =========================
EXCLUDED_FOLDER_NAMES = {"noemailfound", "donotemail"}  # case-insensitive
ATTACH_EXTS = [".pdf"]  # attach only PDFs; set to None to attach all files

FROM_ADDRESS = "Invoices@acmebarricades.com"

# Output behavior
SHOW_SUMMARY = True              # prints a final summary line
PRINT_FAILED_PATHS = True        # prints failed folder paths at end
PRINT_FAILURE_ERRORS = True      # prints error details (folder + reason)
CONTINUE_ON_ERROR = True         # continue to next folder if one fails

# Date formatting in subject
DATE_FORMAT = "%B %Y"  # e.g. "January 2026"

# Column headers in mapping workbook
COL_VERIFIED = "Email Verified? (Yes/No)"
COL_DNI = "DoNotEmail"
COL_FINAL = "FinalFolderName"
COL_CUST = "CustomerName"  # kept for potential fallback (not used for subject now)
COL_DATE = "InvoiceDate"
COL_EMAIL = "Email"
COL_SUBJECT = "Email Subject Line"  # Use this instead of customer name

BODY = (
    "Attached please find your billing for this month. This inbox is not monitored so please contact "
    "the biller noted below if you have any questions. Thank you.\n\n"
    "For our Jacksonville, Orlando, Tallahassee, Panama City, or Pensacola Branches, contact "
    "Stephanie Mata: smata@acmebarricades.com\n\n"
    "For our Miami, West Palm Beach, Tampa, or Ft Myers Branches, contact "
    "Laura Pittman: lpittman@acmebarricades.com\n\n"
    "For our Alabama, Tennessee, Barrier Wall, Guardrail, Perm Signs, or Striping Divisions, contact "
    "Megan Buckley-Finley: mbuckleyfinley@acmebarricades.com\n\n"
    "For all other questions, please contact Kewanna Groman, Billing Team Lead at kgroman@acmebarricades.com\n\n"
    "CORPORATE OFFICE\n"
    "9800 Normandy Blvd\n"
    "Jacksonville, FL 32221\n"
    "Toll-free:  800-373-7704\n"
    "Office:  904-781-1950\n"
    "Fax:  904-781-1921\n"
)

# =========================
# SharePoint/OneDrive logo finder
# =========================
SIGNATURE_IMAGE_FILENAME = "Acme Barricades Logo for Email Signature.jpg"
SIGNATURE_IMAGE_CID = "acme_signature_logo"

def find_signature_logo():
    """
    Search for the signature logo inside OneDrive / SharePoint synced folders.
    Returns absolute path or None.
    """
    possible_roots = []

    # ENV vars that exist on many machines with OneDrive installed
    for var in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        v = os.environ.get(var)
        if v and os.path.isdir(v):
            possible_roots.append(v)

    # Also include user's home directory (recursive search)
    home = os.path.expanduser("~")
    possible_roots.append(home)

    # Deduplicate
    seen = set()
    roots = []
    for r in possible_roots:
        rp = os.path.abspath(r)
        if rp not in seen:
            seen.add(rp)
            roots.append(rp)

    # Search for the file
    for root in roots:
        for dirpath, _, filenames in os.walk(root):
            if SIGNATURE_IMAGE_FILENAME in filenames:
                return os.path.join(dirpath, SIGNATURE_IMAGE_FILENAME)

    return None

# =========================
# Helpers
# =========================
def norm(s: str) -> str:
    return (s or "").strip().lower()

def clean_path(p: str) -> str:
    """Strip whitespace/quotes, normalize, return absolute path."""
    p = (p or "").strip()
    if (p.startswith('"') and p.endswith('"')) or (p.startswith("'") and p.endswith("'")):
        p = p[1:-1].strip()
    return os.path.abspath(os.path.normpath(p))

def parse_recipients_from_email_field(values):
    """Split on ';' or ',' and return de-duped ordered list."""
    seen = set()
    ordered = []
    for val in values:
        if pd.isna(val):
            continue
        s = str(val).strip()
        if not s:
            continue
        parts = re.split(r"[;,]+", s)
        for part in parts:
            addr = part.strip()
            if addr and addr.lower() not in seen:
                seen.add(addr.lower())
                ordered.append(addr)
    return ordered

def collect_attachments(folder_path: str, allowed_exts):
    files = []
    for child in os.scandir(folder_path):
        if child.is_file():
            if allowed_exts is None:
                files.append(child.path)
            else:
                ext = os.path.splitext(child.name)[1].lower()
                if ext in {e.lower() for e in allowed_exts}:
                    files.append(child.path)
    return sorted(files)

def first_non_empty(series):
    for v in series:
        if pd.notna(v) and str(v).strip():
            return v
    return None

def coerce_date(val):
    dt = pd.to_datetime(val, errors="coerce")
    if pd.isna(dt):
        return None
    return dt

def format_date_for_subject(val):
    dt = coerce_date(val)
    if dt is not None:
        return dt.strftime(DATE_FORMAT)
    return str(val).strip()

def resolve_from_account(outlook, from_address: str):
    """
    Try to resolve Outlook Account by SMTP address or display name.
    Returns Account or None.
    """
    try:
        session = outlook.Session
        accounts = session.Accounts
        target = (from_address or "").strip().lower()
        for i in range(1, accounts.Count + 1):
            acc = accounts.Item(i)
            try:
                smtp = (getattr(acc, "SmtpAddress", "") or "").strip().lower()
            except Exception:
                smtp = ""
            display = (getattr(acc, "DisplayName", "") or "").strip().lower()
            if target and (target == smtp or target == display):
                return acc
    except Exception:
        pass
    return None

def stage_attachments_short_paths(files, tag="acme_invoice_mail_staging"):
    staging_dir = os.path.join(tempfile.gettempdir(), tag)

    # SIMPLE FIX: clear the folder so filenames never collide
    if os.path.isdir(staging_dir):
        shutil.rmtree(staging_dir, ignore_errors=True)
    os.makedirs(staging_dir, exist_ok=True)

    staged = []
    for src in files:
        src_abs = os.path.abspath(src)
        base = os.path.basename(src_abs)
        dst = os.path.join(staging_dir, base)

        shutil.copy2(src_abs, dst)
        staged.append(dst)

    return staged

def attach_inline_image(mail, image_path: str, cid: str):
    """
    Attach an image and set Content-ID so it can be referenced inline via <img src="cid:...">
    Works with JPEG/JPG.
    """
    image_path = os.path.abspath(image_path)
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Logo image not found: {image_path}")

    attachment = mail.Attachments.Add(image_path)

    # PR_ATTACH_CONTENT_ID (0x3712001F)
    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
        cid
    )

    # PR_ATTACH_MIME_TAG (0x370E001F) – helpful for JPEG
    ext = os.path.splitext(image_path)[1].lower()
    mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
    try:
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x370E001F",
            mime
        )
    except Exception:
        pass

    # Optional: hide attachment in some clients
    try:
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
            True
        )
    except Exception:
        pass

def html_escape(text: str) -> str:
    if text is None:
        return ""
    return (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;"))

def build_html_body_with_logo_above_corporate(body_text: str, cid: str, show_logo: bool = True) -> str:
    """
    Convert BODY text to HTML and insert an inline logo ABOVE 'CORPORATE OFFICE'.
    If show_logo is False, omit the <img>.
    """
    body_text = body_text or ""
    marker = "CORPORATE OFFICE"
    idx = body_text.find(marker)

    if idx != -1:
        before = body_text[:idx].rstrip()
        after = body_text[idx:].lstrip()
    else:
        before = body_text.rstrip()
        after = ""

    before_html = html_escape(before).replace("\n", "<br>")
    after_html = html_escape(after).replace("\n", "<br>")

    # Inline logo image
    img_html = ""
    if show_logo:
        img_html = (
            f'<div style="margin:12px 0;">'
            f'<img src="cid:{cid}" style="max-width:260px; height:auto;" alt="ACME Logo">'
            f"</div>"
        )

    joiner = "<br><br>" if (img_html or after_html) else ""
    html = before_html + joiner + (img_html if img_html else "") + (after_html if after_html else "")
    return f'<div style="font-family:Calibri, Arial, sans-serif; font-size:11pt;">{html}</div>'

# =========================
# Core send routine
# =========================
def send_grouped_emails_for_path(grouped_path: str, outlook, from_account):
    grouped_path = clean_path(grouped_path)
    mapping_xlsx = os.path.join(grouped_path, "Email_and_PDF_Mapping.xlsx")

    if not os.path.isdir(grouped_path):
        raise FileNotFoundError(f"'Grouped_PDFs' folder not found: {grouped_path}")
    if not os.path.isfile(mapping_xlsx):
        raise FileNotFoundError(f"Mapping file not found: {mapping_xlsx}")

    # Resolve logo once (do NOT attach yet; mail object doesn't exist here)
    logo_abs = find_signature_logo()
    if not logo_abs:
        print("WARNING: Signature image not found. Sending without logo.")

    df = pd.read_excel(mapping_xlsx, engine="openpyxl")

    # Filters: Verified Yes, FinalFolderName present, DoNotEmail blank
    verified_yes = df[COL_VERIFIED].astype(str).str.strip().str.lower() == "yes"
    final_present = df[COL_FINAL].astype(str).str.strip() != ""
    dni_blank = df[COL_DNI].isna() | (df[COL_DNI].astype(str).str.strip() == "")
    df_send = df.loc[verified_yes & final_present & dni_blank].copy()

    # Required columns: use Email Subject Line instead of CustomerName
    missing_cols = [c for c in [COL_FINAL, COL_SUBJECT, COL_DATE, COL_EMAIL] if c not in df_send.columns]
    if missing_cols:
        raise KeyError(f"Required columns missing in mapping file: {missing_cols}")

    # Build per-folder metadata
    folder_meta = {}
    for folder_name, g in df_send.groupby(COL_FINAL):
        folder_key = str(folder_name).strip()
        if norm(folder_key) in EXCLUDED_FOLDER_NAMES:
            continue

        subj_val = first_non_empty(g[COL_SUBJECT])
        date_raw = first_non_empty(g[COL_DATE])
        recipients = parse_recipients_from_email_field(g[COL_EMAIL])

        if not str(subj_val).strip():
            # Subject line is required for subject construction
            raise ValueError(f"Email Subject Line is blank for folder '{folder_key}'.")
        if not str(date_raw).strip():
            raise ValueError(f"InvoiceDate is blank for folder '{folder_key}'.")
        if not recipients:
            # If no recipients, skip this folder silently
            continue

        folder_meta[folder_key] = {
            "recipients": recipients,
            "subject_line": str(subj_val).strip(),
            "date_str": format_date_for_subject(date_raw),
        }

    sent_count = 0
    failed_count = 0

    # Track failures for printing
    failed_folder_errors = []  # list of (folder_path, folder_key, error_str)

    for folder_key, meta in folder_meta.items():
        folder_path = os.path.join(grouped_path, folder_key)

        try:
            if not os.path.isdir(folder_path):
                failed_count += 1
                failed_folder_errors.append((folder_path, folder_key, "Folder not found on disk"))
                continue

            # === Attach ONLY PDFs whose mapping rows are Verified = Yes (using 'source_pdf' column) ===
            # 1) All PDFs present on disk for this folder
            all_pdfs_on_disk = collect_attachments(folder_path, ATTACH_EXTS)
            if not all_pdfs_on_disk:
                failed_count += 1
                failed_folder_errors.append((folder_path, folder_key, "No PDF attachments found"))
                continue

            # 2) Rows in the mapping for this folder (both Yes/No)
            g_all = df[df[COL_FINAL].astype(str).str.strip() == folder_key]

            # 3) Keep only rows where Email Verified? (Yes/No) == "yes"
            g_yes = g_all[g_all[COL_VERIFIED].astype(str).str.strip().str.lower() == "yes"]

            # 4) Build a lookup for PDFs on disk (case-insensitive)
            disk_map = {os.path.basename(p).strip().lower(): p for p in all_pdfs_on_disk}

            # 5) Collect allowed PDFs by matching mapping 'source_pdf' filenames to disk files
            allowed = []
            missing_on_disk = []
            for _, row in g_yes.iterrows():
                pdf_name = str(row.get("source_pdf", "")).strip().lower()
                if not pdf_name:
                    # ignore empty filename cells
                    continue
                # direct match (exact filename)
                if pdf_name in disk_map:
                    allowed.append(disk_map[pdf_name])
                    continue
                # relaxed match: if mapping omitted/changed case/spaces around basename
                # (e.g., sometimes users copy names with stray spaces)
                # We also try a basename match (without extension) for a little resilience.
                mapped_base = os.path.splitext(os.path.basename(pdf_name))[0]
                matched = False
                for disk_name, disk_path in disk_map.items():
                    disk_base = os.path.splitext(os.path.basename(disk_name))[0]
                    if disk_name == pdf_name or disk_base == mapped_base:
                        allowed.append(disk_path)
                        matched = True
                        break
                if not matched:
                    missing_on_disk.append(row.get("source_pdf", ""))

            # 6) If nothing to attach, skip this folder with a helpful reason
            if not allowed:
                msg = "No PDFs marked YES matched on disk for this folder."
                if missing_on_disk:
                    msg += f" Missing from disk (from mapping): {', '.join(map(str, missing_on_disk))}"
                failed_count += 1
                failed_folder_errors.append((folder_path, folder_key, msg))
                continue

            # 7) Stage only the allowed (YES) PDFs
            staged_pdfs = stage_attachments_short_paths(allowed)

            # Subject now uses Email Subject Line instead of CustomerName
            subject = f"{meta['subject_line']} - Billing for {meta['date_str']}"
            recipients = meta["recipients"]

            mail = outlook.CreateItem(0)

            # Sender selection (silent)
            if from_account:
                try:
                    mail.SendUsingAccount = from_account
                except Exception:
                    mail.SentOnBehalfOfName = FROM_ADDRESS
            else:
                mail.SentOnBehalfOfName = FROM_ADDRESS

            mail.To = "; ".join(recipients)
            mail.CC = FROM_ADDRESS
            mail.Subject = subject

            # HTML body + inline logo above "CORPORATE OFFICE"
            mail.HTMLBody = build_html_body_with_logo_above_corporate(
                BODY,
                SIGNATURE_IMAGE_CID,
                show_logo=bool(logo_abs)
            )

            # Attach inline logo only if found
            if logo_abs:
                attach_inline_image(mail, logo_abs, SIGNATURE_IMAGE_CID)

            # Attach PDFs
            for p in staged_pdfs:
                mail.Attachments.Add(p)

            # Send (treat "moved or deleted" as success)
            try:
                mail.Send()
                sent_count += 1
            except Exception as e:
                msg = str(e).lower()
                if "moved or deleted" in msg:
                    sent_count += 1
                else:
                    failed_count += 1
                    failed_folder_errors.append((folder_path, folder_key, repr(e)))
                    if not CONTINUE_ON_ERROR:
                        raise
            finally:
                mail = None

        except Exception as e:
            failed_count += 1
            failed_folder_errors.append((folder_path, folder_key, repr(e)))
            if not CONTINUE_ON_ERROR:
                raise

    return sent_count, failed_count, failed_folder_errors

# =========================
# Main
# =========================
def main():
    outlook = win32.Dispatch("Outlook.Application")
    from_account = resolve_from_account(outlook, FROM_ADDRESS)

    while True:
        grouped = input("\nEnter the path to the 'Grouped_PDFs' folder: ").strip()
        grouped = clean_path(grouped)

        failed_folder_errors = []
        sent_count = 0
        failed_count = 0

        try:
            sent_count, failed_count, failed_folder_errors = send_grouped_emails_for_path(grouped, outlook, from_account)
        except Exception as e:
            # One-line fatal output (still quiet)
            failed_count += 1
            failed_folder_errors.append((grouped, "(RUN)", repr(e)))

        if SHOW_SUMMARY:
            print(f"\nDone. Sent: {sent_count} | Failed: {failed_count}")

        if PRINT_FAILED_PATHS and failed_folder_errors:
            print("\nFailed folder paths:")
            for folder_path, folder_key, _ in failed_folder_errors:
                print(f"- {folder_path}")

        if PRINT_FAILURE_ERRORS and failed_folder_errors:
            print("\nFailurers:")
            for folder_path, folder_key, error_message in failed_folder_errors:
                print(f"- {folder_key}")
                print(f"  Path : {folder_path}")
                print(f"  Error: {error_message}")

        again = input("\nDo you want to send more verified emails to customers? (yes/no): ").strip().lower()
        if again not in ("y", "yes"):
            break

if __name__ == "__main__":
    main()