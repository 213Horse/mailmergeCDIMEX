from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Callable, List, Tuple

import streamlit as st

from send_mail_merge import run_merge


def save_upload(file, suffix: str) -> Path:
    """Persist an uploaded file to a temporary path and return the path."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.read())
        return Path(tmp.name)


def main() -> None:
    st.set_page_config(page_title="Bookmedi Mail Merge", layout="centered")
    st.title("Bookmedi Mail Merge")
    st.caption("Gửi email kèm PDF theo danh sách")

    base = Path(__file__).parent
    default_recipients = base / "recipients.xlsx"
    default_template = base / "template.html"

    # Sidebar: SMTP settings
    st.sidebar.header("SMTP Settings")
    smtp_provider = st.sidebar.selectbox(
        "Provider",
        ["Gmail (STARTTLS)", "Outlook (STARTTLS)", "Custom"],
        index=0,
    )

    host_default = "smtp.gmail.com" if smtp_provider == "Gmail (STARTTLS)" else (
        "smtp.office365.com" if smtp_provider == "Outlook (STARTTLS)" else "smtp.gmail.com"
    )

    smtp_host = st.sidebar.text_input("SMTP Host", value=host_default)
    smtp_port = st.sidebar.number_input("SMTP Port", min_value=1, max_value=65535, value=587)
    use_ssl = st.sidebar.checkbox("Use SSL (SMTPS)", value=False)
    smtp_user = st.sidebar.text_input("SMTP User (email)", value="")
    smtp_pass = st.sidebar.text_input("SMTP Password/App Password", type="password", value="")
    from_name = st.sidebar.text_input("From Name", value="Bookmedi")
    default_subject = st.sidebar.text_input(
        "Default Subject",
        value="Kết quả bài thi Versant level 1 - {{Ten}}",
    )
    dry_run = st.sidebar.checkbox("Dry-run (không gửi thật)", value=True)
    rate_delay = st.sidebar.slider("Delay giữa mỗi email (giây)", 0.0, 10.0, 1.5, 0.5)

    # Main form
    st.subheader("Chọn tệp")
    up_recipients = st.file_uploader("Recipients (.xlsx/.csv)", type=["xlsx", "xls", "csv"], help="Bắt buộc")
    up_template = st.file_uploader("Template HTML", type=["html"], help="Mặc định dùng template.html trong dự án nếu để trống")

    # Or use defaults from project folder
    st.write(":small_blue_diamond: Nếu không upload, ứng dụng sẽ dùng:")
    st.code(str(default_recipients), language="text")
    st.code(str(default_template), language="text")

    log_area = st.empty()
    start = st.button("Gửi Email")

    if start:
        try:
            if up_recipients is not None:
                recipients_path = save_upload(up_recipients, suffix=Path(up_recipients.name).suffix)
            else:
                if not default_recipients.exists():
                    st.error("Chưa chọn recipients và không tìm thấy recipients.xlsx mặc định.")
                    return
                recipients_path = default_recipients

            if up_template is not None:
                template_path = save_upload(up_template, suffix=".html")
            else:
                if not default_template.exists():
                    st.error("Chưa chọn template và không tìm thấy template.html mặc định.")
                    return
                template_path = default_template

            # Progress logs
            session_logs: List[str] = []

            def cb(line: str) -> None:
                session_logs.append(line)
                log_area.text("\n".join(session_logs))

            summary = run_merge(
                recipients=str(recipients_path),
                template=str(template_path),
                smtp_host=smtp_host,
                smtp_port=int(smtp_port),
                smtp_user=smtp_user,
                smtp_pass=smtp_pass,
                from_name=from_name,
                default_subject=default_subject,
                rate_delay=float(rate_delay),
                dry_run=bool(dry_run),
                use_ssl=bool(use_ssl),
                progress_callback=cb,
            )

            st.success(f"Hoàn tất. Sent={summary['sent']}, Failed={summary['failed']}")
            if summary["errors"]:
                with st.expander("Xem lỗi"):
                    for em, err in summary["errors"]:
                        st.write(f"- {em}: {err}")
        except Exception as exc:
            st.error(str(exc))


if __name__ == "__main__":
    main()


