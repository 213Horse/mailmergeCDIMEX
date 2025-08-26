from __future__ import annotations

# ==== Giới hạn thread để không ăn 100% CPU (đặt sớm nhất) ====
import os
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("MKL_NUM_THREADS", "1")
os.environ.setdefault("NUMEXPR_NUM_THREADS", "1")

import time
import tempfile
from pathlib import Path
from typing import List
import shutil
import zipfile
import fcntl  # dùng lock file trên Linux

import streamlit as st
try:
    # Optional rich text editor
    from streamlit_ckeditor import st_ckeditor  # type: ignore
except Exception:
    st_ckeditor = None  # fallback later
try:
    from streamlit_quill import st_quill  # type: ignore
except Exception:
    st_quill = None

from send_mail_merge import run_merge


# ========== Tiện ích chung ==========
def _safe_rerun() -> None:
    """Call Streamlit rerun API across versions."""
    rerun_fn = getattr(st, "rerun", None)
    if callable(rerun_fn):
        rerun_fn()
        return
    exp_rerun_fn = getattr(st, "experimental_rerun", None)
    if callable(exp_rerun_fn):
        exp_rerun_fn()
        return


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


class ThrottledLogger:
    """Giảm tần suất cập nhật UI để đỡ tốn CPU.

    - batch_size: gom N dòng rồi mới vẽ
    - min_interval: ít nhất mỗi X giây mới vẽ 1 lần
    """
    def __init__(self, placeholder, batch_size: int = 12, min_interval: float = 0.25):
        self.placeholder = placeholder
        self.batch_size = batch_size
        self.min_interval = min_interval
        self._buf: List[str] = []
        self._all: List[str] = []
        self._last_flush = 0.0

    def __call__(self, line: str) -> None:
        self._buf.append(line)
        self._all.append(line)
        now = time.time()
        if len(self._buf) >= self.batch_size or (now - self._last_flush) >= self.min_interval:
            # chỉ render một lần cho cả batch
            self.placeholder.text("\n".join(self._all[-800:]))  # tránh render quá dài
            self._buf.clear()
            self._last_flush = now

    def flush(self) -> None:
        if self._buf:
            self.placeholder.text("\n".join(self._all[-800:]))
            self._buf.clear()
            self._last_flush = time.time()


# ========== File Manager ==========
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
        if st.button("Back", disabled=fm_cwd == fm_root, key="fm_back"):
            parent = fm_cwd.parent
            if _is_safe_relative_path(fm_root, parent):
                st.session_state["fm_cwd"] = str(parent)
                _safe_rerun()
    with cols[2]:
        if st.button("Làm mới", key="fm_refresh"):
            _safe_rerun()
    with cols[3]:
        allow_delete = st.checkbox("Bật xoá", value=False, key="fm_allow_delete")

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
                        _safe_rerun()
                    except FileExistsError:
                        st.warning("Thư mục đã tồn tại.")
                    except Exception as exc:
                        st.error(f"Không thể tạo thư mục: {exc}")

    # Upload files
    uploads = st.file_uploader("Upload file vào thư mục hiện tại", accept_multiple_files=True, key="fm_uploader")
    if uploads:
        if st.button("Lưu các file upload", key="fm_save_uploads"):
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
            _safe_rerun()

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
                        _safe_rerun()
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
                            _safe_rerun()
                        except Exception as exc:
                            st.error(f"Không thể xoá thư mục: {exc}")
            else:
                if st.button("Xoá", key=f"delf_{e.name}", disabled=not allow_delete):
                    if _is_safe_relative_path(fm_root, e):
                        try:
                            e.unlink(missing_ok=False)
                            st.success(f"Đã xoá tệp: {e.name}")
                            _safe_rerun()
                        except Exception as exc:
                            st.error(f"Không thể xoá tệp: {exc}")


# ========== App chính ==========
def main() -> None:
    st.set_page_config(page_title="Cdimex Mail Merge", layout="centered")
    st.title("Cdimex Mail Merge")
    st.caption("Gửi email kèm PDF theo danh sách")

    base = Path(__file__).parent
    upload_dir = base / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    default_recipients = base / "recipients.xlsx"
    default_template = base / "template.html"

    # Trạng thái chạy để chặn bấm nhiều lần
    if "running" not in st.session_state:
        st.session_state["running"] = False

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
        up_recipients = st.file_uploader("Nhập file danh sách khách hàng (.xlsx/.csv)", type=["xlsx", "xls", "csv"], help="Bắt buộc", key="rec_upl")

        MODE_FILE = "Dùng file HTML"
        MODE_EDITOR = "Soạn trực tiếp (WYSIWYG)"
        mode = st.radio(
            "Cách nhập nội dung email",
            [MODE_FILE, MODE_EDITOR],
            horizontal=True,
            index=1,
            key="content_mode",
        )

        html_content: str | None = None
        up_template = None
        if mode == MODE_FILE:
            up_template = st.file_uploader(
                "Template HTML",
                type=["html"],
                help="Mặc định dùng template.html trong dự án nếu để trống",
                key="tpl_upl",
            )
        else:
            st.caption("Bạn có thể gõ nội dung và dùng token như {{Ten}}, {{Email}} ...")
            default_html = ""
            try:
                if default_template.exists():
                    default_html = default_template.read_text(encoding="utf-8")
            except Exception:
                default_html = ""

            editor_key = "editor_html"
            if st_ckeditor is not None:
                # CKEditor: full WYSIWYG, users don't need to know HTML
                html_content = st_ckeditor(
                    default_html,
                    key=editor_key,
                    height=320,
                )
            elif st_quill is not None:
                # Quill editor: also WYSIWYG, return HTML for sending
                editor_key = "editor_quill_html"
                html_content = st_quill(
                    html=True,
                    placeholder="Soạn nội dung email...",
                    key=editor_key,
                )
            else:
                # Last resort: textarea (not ideal but keeps app usable)
                editor_key = "editor_html_fallback"
                html_content = st.text_area(
                    "Nội dung (HTML)",
                    value=default_html,
                    height=320,
                    key=editor_key,
                )

            with st.expander("Chèn token nhanh", expanded=False):
                col_t = st.columns(3)
                tokens = ["{{Ten}}", "{{Email}}", "{{NgayGui}}"]
                for i, tk in enumerate(tokens):
                    if col_t[i].button(tk, key=f"ins_{tk}"):
                        try:
                            cur = st.session_state.get(editor_key, "") or ""
                            st.session_state[editor_key] = f"{cur}{tk}"
                            _safe_rerun()
                        except Exception:
                            pass

        st.subheader("File đính kèm")
        st.caption("Nếu recipients chỉ chứa tên file (ví dụ: A.pdf), hãy upload thư mục ZIP chứa các PDF hoặc chọn thư mục gốc.")
        zip_upload = st.file_uploader("Upload ZIP chứa các PDF (tùy chọn)", type=["zip"], accept_multiple_files=False, key="zip_upl")
        pdf_uploads = st.file_uploader("Upload nhiều PDF (tùy chọn)", type=["pdf"], accept_multiple_files=True, key="pdfs_upl")
        base_dir_text = st.text_input("Base directory (tùy chọn)", value=str(upload_dir), key="base_dir_txt")

        # Or use defaults from project folder
        st.write(":small_blue_diamond: Nếu không upload, ứng dụng sẽ dùng:")
        st.code(str(default_recipients), language="text")
        st.code(str(default_template), language="text")

        log_area = st.empty()
        throttled_log = ThrottledLogger(log_area)

        start = st.button("Gửi Email", disabled=st.session_state["running"])
        if start:
            # ====== Lock: ngăn chạy trùng tiến trình ======
            st.session_state["running"] = True
            lock_path = Path("/tmp/bookmedi_mailmerge.lock")
            try:
                with open(lock_path, "w") as lock_file:
                    try:
                        fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
                    except BlockingIOError:
                        st.warning("Job khác đang chạy. Vui lòng đợi hoàn tất rồi thử lại.")
                        st.session_state["running"] = False
                        return

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
                            # Giới hạn kích thước zip để tránh out-of-memory
                            if zip_upload.size and zip_upload.size > 100 * 1024 * 1024:
                                st.error(f"ZIP quá lớn ({_human_size(zip_upload.size)}). Vui lòng chia nhỏ (< 100MB).")
                                st.session_state["running"] = False
                                return
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
                                st.session_state["running"] = False
                                return
                            recipients_path = default_recipients

                        if mode == MODE_FILE:
                            if up_template is not None:
                                template_path = save_upload(up_template, suffix=".html")
                            else:
                                if not default_template.exists():
                                    st.error("Chưa chọn template và không tìm thấy template.html mặc định.")
                                    st.session_state["running"] = False
                                    return
                                template_path = default_template
                        else:
                            # Soạn trực tiếp: ghi ra file tạm để tái sử dụng luồng cũ
                            content_to_use = (html_content or "").strip()
                            if not content_to_use:
                                st.error("Nội dung email đang trống.")
                                st.session_state["running"] = False
                                return
                            # Lưu trong uploads/ để các đường dẫn tương đối trong HTML có thể tham chiếu tới tệp trong dự án
                            try:
                                editor_tpl = upload_dir / "_editor_template.html"
                                editor_tpl.write_text(content_to_use, encoding="utf-8")
                                template_path = editor_tpl
                            except Exception:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as tmp_html:
                                    tmp_html.write(content_to_use)
                                    template_path = Path(tmp_html.name)

                        # Thông báo ban đầu
                        if saved_files:
                            throttled_log(f"Đã lưu {len(saved_files)} PDF vào: {upload_dir}")
                        elif zip_upload is not None:
                            throttled_log(f"Đã giải nén ZIP vào: {upload_dir}")

                        # ==== Chạy merge với callback đã throttle ====
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
                            progress_callback=throttled_log,
                        )

                        throttled_log.flush()
                        st.success(f"Hoàn tất. Sent={summary['sent']}, Failed={summary['failed']}")
                        if summary.get("errors"):
                            with st.expander("Xem lỗi"):
                                for em, err in summary["errors"]:
                                    st.write(f"- {em}: {err}")
                    finally:
                        try:
                            fcntl.flock(lock_file, fcntl.LOCK_UN)
                        except Exception:
                            pass
            except Exception as exc:
                st.error(str(exc))
            finally:
                st.session_state["running"] = False

    with tab_files:
        render_file_manager(base)


if __name__ == "__main__":
    main()
