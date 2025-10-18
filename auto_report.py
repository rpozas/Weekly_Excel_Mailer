
import os
import smtplib
import ssl
import mimetypes
import traceback
from datetime import datetime
from email.message import EmailMessage

import pandas as pd
import yaml

# -------------- Helpers --------------

def load_config(path: str = "config.yaml") -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def ensure_parent_dir(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)

def log(msg: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {msg}", flush=True)

# -------------- Business Logic (EDIT HERE) --------------

def custom_update(input_xlsx: str, output_xlsx: str) -> None:
    """
    Edit this function to implement your actual update rules.
    Default behavior:
    - Reads the first sheet to a DataFrame (if present)
    - Adds/updates a 'Last_Updated' cell in a small summary sheet
    - Saves to output_xlsx
    """
    # Read first sheet if available
    xls = pd.ExcelFile(input_xlsx)
    first_sheet = xls.sheet_names[0] if xls.sheet_names else None

    if first_sheet:
        df = pd.read_excel(input_xlsx, sheet_name=first_sheet)
    else:
        df = pd.DataFrame()

    # Example transformation: no-op (pass-through)
    # TODO: replace with your real logic, e.g., pulling fresh data, filtering, aggregations, etc.

    # Write out with a small "Summary" tab
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        if not df.empty:
            df.to_excel(writer, sheet_name=first_sheet, index=False)
        # Summary / metadata
        meta = pd.DataFrame({
            "Field": ["Last_Updated", "Source_File", "Rows_In_First_Sheet"],
            "Value": [datetime.now().isoformat(timespec="seconds"), os.path.basename(input_xlsx), len(df)]
        })
        meta.to_excel(writer, sheet_name="Summary", index=False)

# -------------- Email --------------

def attach_file(msg: EmailMessage, file_path: str) -> None:
    ctype, encoding = mimetypes.guess_type(file_path)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)

    with open(file_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(file_path))

def send_email(cfg: dict, attachment_path: str) -> None:
    smtp = cfg["email"]["smtp"]
    msg = EmailMessage()
    msg["Subject"] = cfg["email"]["subject"]
    msg["From"] = cfg["email"]["from"]
    msg["To"] = ", ".join(cfg["email"]["to"])
    if "cc" in cfg["email"] and cfg["email"]["cc"]:
        msg["Cc"] = ", ".join(cfg["email"]["cc"])
    if "bcc" in cfg["email"] and cfg["email"]["bcc"]:
        # For BCC we don't add a header; just include in send recipients below
        pass
    body = cfg["email"].get("body", "")
    msg.set_content(body)

    # Attachment
    if attachment_path:
        attach_file(msg, attachment_path)

    # Resolve recipients including CC/BCC
    recipients = list(cfg["email"]["to"]) + cfg["email"].get("cc", []) + cfg["email"].get("bcc", [])

    if smtp.get("use_tls", True):
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp["host"], smtp["port"]) as server:
            server.starttls(context=context)
            server.login(smtp["username"], smtp["password"])
            server.send_message(msg, to_addrs=recipients)
    else:
        with smtplib.SMTP_SSL(smtp["host"], smtp["port"]) as server:
            server.login(smtp["username"], smtp["password"])
            server.send_message(msg, to_addrs=recipients)

# -------------- Main --------------

def main():
    try:
        cfg = load_config("config.yaml")

        input_xlsx = cfg["files"]["input_xlsx"]
        ts = datetime.now().strftime("%Y%m%d")
        out_dir = cfg["files"].get("output_dir", "out")
        ensure_parent_dir(out_dir + os.sep)
        output_xlsx = os.path.join(out_dir, cfg["files"].get("output_name_pattern", f"report_{ts}.xlsx").format(date=ts))

        log(f"Reading: {input_xlsx}")
        log(f"Writing: {output_xlsx}")
        custom_update(input_xlsx, output_xlsx)

        if cfg.get("email", {}).get("enabled", True):
            log("Sending email...")
            send_email(cfg, output_xlsx)
            log("Email sent.")
        else:
            log("Email disabled in config; skipping send.")

        log("Done.")
    except Exception as e:
        log("ERROR occurred: " + str(e))
        traceback.print_exc()
        raise

if __name__ == "__main__":
    main()
