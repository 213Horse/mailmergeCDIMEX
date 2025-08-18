import argparse
import os
import smtplib
import ssl
import time
import mimetypes
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr, make_msgid
import pandas as pd
from pathlib import Path

REQUIRED_COLS = ["Email", "Ten", "FilePDF"]
OPTIONAL_COLS = ["Subject", "CC", "BCC"]

def normalize_field(value: object) -> str:
    try:
        import pandas as pd  # local import for script
        if isinstance(value, float) and value != value:  # NaN check without pandas
            return ""
        if 'pandas' in str(type(value)).lower():
            # Fallback if a pandas NA type sneaks in; convert to empty
            try:
                if pd.isna(value):
                    return ""
            except Exception:
                pass
        s = str(value).strip()
        if s.lower() == "nan":
            return ""
        return s
    except Exception:
        s = str(value).strip()
        return "" if s.lower() == "nan" else s

def load_recipients(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path)
    elif ext in [".csv"]:
        df = pd.read_csv(path)
    else:
        raise ValueError("Chỉ hỗ trợ .xlsx, .xls, .csv")
    for col in REQUIRED_COLS:
        if col not in df.columns:
            raise ValueError(f"Thiếu cột bắt buộc: {col}")
    # Ensure optional columns exist
    for col in OPTIONAL_COLS:
        if col not in df.columns:
            df[col] = ""
    return df

def render_template(html: str, mapping: dict) -> str:
    # simple {{key}} replacement
    out = html
    for k, v in mapping.items():
        out = out.replace("{{" + k + "}}", str(v))
    return out

def build_message(sender_name, sender_email, to_email, cc, bcc, subject, html_body, text_fallback=None):
    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr((sender_name, sender_email)) if sender_name else sender_email
    msg["To"] = to_email
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject
    msg["Message-ID"] = make_msgid()
    # Text fallback
    if not text_fallback:
        text_fallback = "Xin chào,\n\nVui lòng xem nội dung email ở dạng HTML hoặc file đính kèm.\n\nTrân trọng."
    msg.attach(MIMEText(text_fallback, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg

def attach_file(msg, file_path):
    fpath = Path(file_path)
    if not fpath.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {fpath}")
    ctype, encoding = mimetypes.guess_type(str(fpath))
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(fpath, "rb") as f:
        part = MIMEApplication(f.read(), _subtype=subtype)
    part.add_header("Content-Disposition", "attachment", filename=fpath.name)
    msg.attach(part)

def send_email_smtp(host, port, user, password, use_starttls, msg, to_email, cc="", bcc="", dry_run=False):
    recipients = [to_email]
    if cc:
        recipients += [e.strip() for e in cc.split(",") if e.strip()]
    if bcc:
        recipients += [e.strip() for e in bcc.split(",") if e.strip()]

    if dry_run:
        print(f"[DRY-RUN] Would send to: {recipients}")
        return

    if use_starttls:
        context = ssl.create_default_context()
        with smtplib.SMTP(host, port) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(user, password)
            server.sendmail(msg["From"], recipients, msg.as_string())
    else:
        with smtplib.SMTP_SSL(host, port) as server:
            server.login(user, password)
            server.sendmail(msg["From"], recipients, msg.as_string())

def main():
    parser = argparse.ArgumentParser(description="Mail Merge kèm PDF per-recipient (local).")
    parser.add_argument("--recipients", required=True, help="Đường dẫn recipients .xlsx/.csv")
    parser.add_argument("--template", required=True, help="Đường dẫn template HTML")
    parser.add_argument("--smtp-host", required=True, help="SMTP host (vd: smtp.gmail.com hoặc smtp.office365.com)")
    parser.add_argument("--smtp-port", type=int, default=587, help="SMTP port (Gmail/Office365 STARTTLS = 587)")
    parser.add_argument("--smtp-user", required=True, help="SMTP username (email)")
    parser.add_argument("--smtp-pass", required=True, help="SMTP password (Gmail dùng App Password)")
    parser.add_argument("--from-name", default="", help="Tên hiển thị người gửi (optional)")
    parser.add_argument("--default-subject", default="Kết quả bài thi Versant level 1 - {{Ten}}", help="Subject mặc định nếu cột Subject trống")
    parser.add_argument("--rate-delay", type=float, default=2.0, help="Delay (giây) giữa mỗi email để tránh bị giới hạn")
    parser.add_argument("--dry-run", action="store_true", help="Chạy thử: không gửi email thật")
    parser.add_argument("--use-ssl", action="store_true", help="Dùng SMTPS (SSL) thay vì STARTTLS")
    args = parser.parse_args()

    rec_path = Path(args.recipients)
    tpl_path = Path(args.template)

    df = load_recipients(rec_path)
    html = tpl_path.read_text(encoding="utf-8")

    sent, failed = 0, 0
    errors = []

    for i, row in df.iterrows():
        email = normalize_field(row["Email"]) 
        ten = normalize_field(row["Ten"]) 
        fpdf = normalize_field(row["FilePDF"]) 
        cc = normalize_field(row.get("CC", ""))
        bcc = normalize_field(row.get("BCC", ""))
        subj_tpl = normalize_field(row.get("Subject", "")) or args.default_subject

        tokens = {
            "Ten": ten,
            "Email": email,
            "NgayGui": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        subject = render_template(subj_tpl, tokens)
        body_html = render_template(html, tokens)

        try:
            msg = build_message(args.from_name, args.smtp_user, email, cc, bcc, subject, body_html)
            attach_file(msg, fpdf)

            # Decide SSL vs STARTTLS
            use_starttls = not args.use_ssl
            send_email_smtp(
                host=args.smtp_host,
                port=args.smtp_port,
                user=args.smtp_user,
                password=args.smtp_pass,
                use_starttls=use_starttls,
                msg=msg,
                to_email=email,
                cc=cc,
                bcc=bcc,
                dry_run=args.dry_run,
            )
            sent += 1
            print(f"[OK] {email}")
        except Exception as e:
            failed += 1
            errors.append((email, str(e)))
            print(f"[ERR] {email} -> {e}")

        time.sleep(max(0.0, args.rate_delay))

    print(f"\nDone. Sent={sent}, Failed={failed}")
    if errors:
        print("Errors:")
        for em, err in errors:
            print(f" - {em}: {err}")

if __name__ == "__main__":
    main()