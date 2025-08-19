from __future__ import annotations

import tempfile
from pathlib import Path
from typing import List
import shutil

import streamlit as st
import zipfile

from send_mail_merge import run_merge


def save_upload(file, suffix: str) -> Path:
    """Persist an uploaded file to a temporary path and return the path."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(file.read())
        return Path(tmp.name)


def _is_safe_relative_path(base_dir: Path, candidate: Path) -> bool:
    """Return True if candidate resolves under base_dir; False otherwise."""
    try:
        return candidate.resolve().is_relative_to(base_dir.resolve())  # type: ignore[attr-defined]
    except AttributeError:
        # Python < 3.9 fallback
        try:
            candidate.resolve().relative_to(base_dir.resolve())
            return True
        except Exception:
            return False
    except Exception:
        return False


def _human_size(num_bytes: int) -> str:
    units = ["B", "KB", "MB", "GB", "TB"]
    size = float(num_bytes)
    unit_idx = 0
    while size >= 1024.0 and unit_idx < len(units) - 1:
        size /= 1024.0
        unit_idx += 1
    if unit_idx == 0:
        return f"{int(size)} {units[unit_idx]}"
    return f"{size:.1f} {units[unit_idx]}"


def render_file_manager(root_dir: Path) -> None:
    st.header("Quản lý tệp & thư mục")
    st.caption("Thao tác trong phạm vi thư mục dự án để an toàn.")

    # Session state
    if "fm_root" not in st.session_state:
        st.session_state["fm_root"] = str(root_dir.resolve())
    if "fm_cwd" not in st.session_state:
        st.session_state["fm_cwd"] = st.session_state["fm_root"]

    project_base = root_dir.resolve()
    fm_root = Path(st.session_state["fm_root"]).resolve()
    if not _is_safe_relative_path(project_base, fm_root):
        fm_root = project_base
        st.session_state["fm_root"] = str(fm_root)
        st.session_state["fm_cwd"] = str(fm_root)

    try:
        fm_root.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        st.error(f"Không thể tạo thư mục gốc: {exc}")
        return

    fm_cwd = Path(st.session_state["fm_cwd"]).resolve()
    if not _is_safe_relative_path(fm_root, fm_cwd):
        fm_cwd = fm_root
        st.session_state["fm_cwd"] = str(fm_cwd)

    cols = st.columns([5, 2, 2, 3])
    with cols[0]:
        st.write(f"Đang ở: {fm_cwd}")
    with cols[1]:
        if st.button("Lên trên", disabled=fm_cwd == fm_root):
            parent = fm_cwd.parent
            if _is_safe_relative_path(fm_root, parent):
                st.session_state["fm_cwd"] = str(parent)
                st.experimental_rerun()
    with cols[2]:
        if st.button("Làm mới"):
            st.experimental_rerun()
    with cols[3]:
        allow_delete = st.checkbox("Bật xoá", value=False)

    # Create folder form
    with st.form("fm_create_folder", clear_on_submit=True):
        new_folder = st.text_input("Tạo thư mục mới", value="")
        submit_folder = st.form_submit_button("Tạo thư mục")
        if submit_folder:
            name = new_folder.strip().strip("/\\")
            if not name:
                st.warning("Tên thư mục không được rỗng.")
            else:
                target = (fm_cwd / name)
                if not _is_safe_relative_path(fm_root, target):
                    st.error("Đường dẫn không hợp lệ.")
                else:
                    try:
                        target.mkdir(exist_ok=False)
                        st.success(f"Đã tạo: {target.name}")
                        st.experimental_rerun()
                    except FileExistsError:
                        st.warning("Thư mục đã tồn tại.")
                    except Exception as exc:
                        st.error(f"Không thể tạo thư mục: {exc}")

    # Upload files
    uploads = st.file_uploader("Upload file vào thư mục hiện tại", accept_multiple_files=True)
    if uploads:
        if st.button("Lưu các file upload"):
            saved = 0
            for uf in uploads:
                dest = fm_cwd / Path(uf.name).name
                if not _is_safe_relative_path(fm_root, dest):
                    st.error(f"Bỏ qua tệp không hợp lệ: {uf.name}")
                    continue
                try:
                    with open(dest, "wb") as f:
                        f.write(uf.getbuffer())
                    saved += 1
                except Exception as exc:
                    st.error(f"Không thể lưu {uf.name}: {exc}")
            st.success(f"Đã lưu {saved} tệp vào {fm_cwd}")
            st.experimental_rerun()

    st.divider()

    # List entries
    try:
        entries = list(fm_cwd.iterdir())
    except Exception as exc:
        st.error(f"Không thể liệt kê thư mục: {exc}")
        return
    entries.sort(key=lambda p: (p.is_file(), p.name.lower()))

    for e in entries:
        icon = "📁" if e.is_dir() else "📄"
        row = st.columns([6, 2, 2, 2])
        with row[0]:
            st.write(f"{icon} {e.name}")
        with row[1]:
            try:
                size = _human_size(e.stat().st_size) if e.is_file() else "—"
            except Exception:
                size = "—"
            st.write(size)
        with row[2]:
            if e.is_dir():
                if st.button("Mở", key=f"open_{e.name}"):
                    if _is_safe_relative_path(fm_root, e):
                        st.session_state["fm_cwd"] = str(e.resolve())
                        st.experimental_rerun()
            else:
                try:
                    data = e.read_bytes()
                except Exception:
                    data = b""
                st.download_button("Tải", data=data, file_name=e.name, mime="application/octet-stream", key=f"dl_{e.name}")
        with row[3]:
            if e.is_dir():
                if st.button("Xoá", key=f"del_{e.name}", disabled=not allow_delete):
                    if _is_safe_relative_path(fm_root, e):
                        try:
                            shutil.rmtree(e)
                            st.success(f"Đã xoá thư mục: {e.name}")
                            st.experimental_rerun()
                        except Exception as exc:
                            st.error(f"Không thể xoá thư mục: {exc}")
            else:
                if st.button("Xoá", key=f"delf_{e.name}", disabled=not allow_delete):
                    if _is_safe_relative_path(fm_root, e):
                        try:
                            e.unlink(missing_ok=False)
                            st.success(f"Đã xoá tệp: {e.name}")
                            st.experimental_rerun()
                        except Exception as exc:
                            st.error(f"Không thể xoá tệp: {exc}")

def main() -> None:
    st.set_page_config(page_title="Bookmedi Mail Merge", layout="centered")
    st.title("Bookmedi Mail Merge")
    st.caption("Gửi email kèm PDF theo danh sách")

    base = Path(__file__).parent
    upload_dir = base / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    default_recipients = base / "recipients.xlsx"
    default_template = base / "template.html"

    tab_send, tab_files = st.tabs(["Gửi Email", "Quản lý tệp & thư mục"])

    with tab_send:
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

        st.subheader("File đính kèm")
        st.caption("Nếu recipients chỉ chứa tên file (ví dụ: A.pdf), hãy upload thư mục ZIP chứa các PDF hoặc chọn thư mục gốc.")
        zip_upload = st.file_uploader("Upload ZIP chứa các PDF (tùy chọn)", type=["zip"], accept_multiple_files=False)
        pdf_uploads = st.file_uploader("Upload nhiều PDF (tùy chọn)", type=["pdf"], accept_multiple_files=True)
        base_dir_text = st.text_input("Base directory (tùy chọn)", value=str(upload_dir))

        # Or use defaults from project folder
        st.write(":small_blue_diamond: Nếu không upload, ứng dụng sẽ dùng:")
        st.code(str(default_recipients), language="text")
        st.code(str(default_template), language="text")

        log_area = st.empty()
        start = st.button("Gửi Email")

        if start:
            try:
                # Persist uploaded PDFs/ZIP to server under uploads/
                saved_files: List[Path] = []
                if pdf_uploads:
                    for up in pdf_uploads:
                        dest = upload_dir / Path(up.name).name
                        with open(dest, "wb") as f:
                            f.write(up.getbuffer())
                        saved_files.append(dest)

                if zip_upload is not None:
                    tmp_zip = save_upload(zip_upload, suffix=".zip")
                    try:
                        with zipfile.ZipFile(tmp_zip, "r") as zf:
                            zf.extractall(upload_dir)
                    finally:
                        try:
                            tmp_zip.unlink(missing_ok=True)  # type: ignore[arg-type]
                        except Exception:
                            pass

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

                if saved_files:
                    cb(f"Đã lưu {len(saved_files)} PDF vào: {upload_dir}")
                elif zip_upload is not None:
                    cb(f"Đã giải nén ZIP vào: {upload_dir}")

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
                    base_dir=(base_dir_text or str(upload_dir)),
                    progress_callback=cb,
                )

                st.success(f"Hoàn tất. Sent={summary['sent']}, Failed={summary['failed']}")
                if summary["errors"]:
                    with st.expander("Xem lỗi"):
                        for em, err in summary["errors"]:
                            st.write(f"- {em}: {err}")
            except Exception as exc:
                st.error(str(exc))

    with tab_files:
        render_file_manager(base)


if __name__ == "__main__":
    main()


