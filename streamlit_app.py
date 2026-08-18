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
from typing import List, Tuple, Dict
import shutil
import zipfile
import fcntl  # dùng lock file trên Linux
import base64
import mimetypes
import re

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

from send_mail_merge import run_merge, render_template


# ========== Tiện ích chung ==========
def _env_str(name: str, default: str = "") -> str:
    v = os.getenv(name)
    return default if v is None else v


def _env_int(name: str, default: int) -> int:
    v = os.getenv(name)
    if v is None or v.strip() == "":
        return default
    try:
        return int(v)
    except Exception:
        return default


def _env_float(name: str, default: float) -> float:
    v = os.getenv(name)
    if v is None or v.strip() == "":
        return default
    try:
        return float(v)
    except Exception:
        return default


def _env_bool(name: str, default: bool) -> bool:
    v = os.getenv(name)
    if v is None or v.strip() == "":
        return default
    s = v.strip().lower()
    if s in {"1", "true", "yes", "y", "on"}:
        return True
    if s in {"0", "false", "no", "n", "off"}:
        return False
    return default


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


def _extract_body_html(html: str) -> str:
    """Return inner HTML of <body> if present, else return original string.

    Users may paste full HTML documents into header/footer editors; for email composition
    we only want the body portion to avoid nested <html>/<body>.
    """
    if not html:
        return ""
    m = re.search(r"<body\b[^>]*>(?P<body>[\s\S]*?)</body>", html, flags=re.IGNORECASE)
    if m:
        return (m.group("body") or "").strip()
    return html.strip()


def _extract_style_blocks(html: str) -> str:
    if not html:
        return ""
    blocks = re.findall(r"<style\b[^>]*>[\s\S]*?</style>", html, flags=re.IGNORECASE)
    return "\n".join(blocks).strip()


def _looks_empty_rich_html(html: str | None) -> bool:
    """Heuristic for WYSIWYG editors that return '<p><br></p>' etc."""
    if html is None:
        return True
    s = str(html).strip()
    if not s:
        return True
    # remove common empty constructs
    s = s.replace("&nbsp;", " ")
    # strip tags
    s = re.sub(r"<[^>]+>", "", s)
    if s.strip() == "":
        return True
    return False


def _compose_email_html(header_html: str, body_html: str, footer_html: str) -> str:
    """Compose a single HTML document for sending / preview."""
    header_body = _extract_body_html(header_html)
    body_body = _extract_body_html(body_html)
    footer_body = _extract_body_html(footer_html)

    style_blocks = "\n".join(
        b
        for b in [
            _extract_style_blocks(header_html),
            _extract_style_blocks(body_html),
            _extract_style_blocks(footer_html),
        ]
        if b
    ).strip()

    # Email-friendly wrapper: centered 600px content
    parts: List[str] = [
        "<!doctype html>",
        "<html>",
        "<head>",
        '  <meta charset="utf-8"/>',
        '  <meta name="viewport" content="width=device-width, initial-scale=1"/>',
    ]
    if style_blocks:
        parts.append(style_blocks)
    parts += [
        "</head>",
        '<body style="margin:0;padding:0;">',
        '  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" style="background:#ffffff;">',
        "    <tr>",
        '      <td align="center" style="padding:0;">',
        '        <table width="600" border="0" cellspacing="0" cellpadding="0" style="max-width:600px;width:100%;">',
        "          <tr><td>",
        header_body,
        body_body,
        footer_body,
        "          </td></tr>",
        "        </table>",
        "      </td>",
        "    </tr>",
        "  </table>",
        "</body>",
        "</html>",
    ]
    return "\n".join([p for p in parts if p is not None])


def _file_to_data_url(path: Path) -> str | None:
    try:
        mime, _ = mimetypes.guess_type(str(path))
        if not mime:
            ext = path.suffix.lower().lstrip(".")
            if ext in {"png", "jpg", "jpeg", "gif", "webp", "bmp"}:
                mime = f"image/{'jpeg' if ext in {'jpg', 'jpeg'} else ext}"
        if not mime or not mime.startswith("image/"):
            return None
        data = path.read_bytes()
        b64 = base64.b64encode(data).decode("ascii")
        return f"data:{mime};base64,{b64}"
    except Exception:
        return None


def _preview_html_with_embedded_images(html: str, base_dir: Path, cid_logo_filename: str = "logomedi.png") -> str:
    """For Streamlit preview only: replace local/cid logo images with data URLs."""
    if not html:
        return ""

    project_base = Path(__file__).parent.resolve()
    logo_path = project_base / Path(cid_logo_filename).name
    logo_data_url = _file_to_data_url(logo_path) if logo_path.exists() else None
    if logo_data_url:
        html = html.replace("cid:bookmedi_logo", logo_data_url)

    # Replace <img src="relative/or/local"> with base64 for preview
    pattern = re.compile(r'(<img\b[^>]*?\bsrc=)(["\'])([^"\']+)(\2)', re.IGNORECASE | re.DOTALL)

    def _replace(m: re.Match) -> str:
        prefix, quote, src, suffix_quote = m.group(1), m.group(2), (m.group(3) or "").strip(), m.group(4)
        if src.startswith(("http://", "https://", "data:", "cid:")):
            return m.group(0)
        candidate = Path(src)
        resolved = candidate
        if not candidate.is_absolute():
            # Prefer base_dir (uploads/) then project root
            resolved = (base_dir / candidate)
            if not resolved.exists():
                alt = project_base / candidate
                if alt.exists():
                    resolved = alt
        if not resolved.exists():
            return m.group(0)
        data_url = _file_to_data_url(resolved)
        if not data_url:
            return m.group(0)
        return f"{prefix}{quote}{data_url}{suffix_quote}"

    return pattern.sub(_replace, html)


def _load_preview_tokens(uploaded_recipients) -> Dict[str, str]:
    """Build preview token mapping from uploaded recipients (first row) if possible."""
    tokens: Dict[str, str] = {
        "Ten": "Nguyễn Văn A",
        "Email": "nguyenvana@example.com",
        "Code": "ABC123",
        "NgayGui": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    if uploaded_recipients is None:
        return tokens
    try:
        name = getattr(uploaded_recipients, "name", "") or ""
        suffix = Path(name).suffix.lower()
        uploaded_recipients.seek(0)
        if suffix in {".xlsx", ".xls"}:
            import pandas as pd  # local import
            df = pd.read_excel(uploaded_recipients)
        elif suffix == ".csv":
            import pandas as pd  # local import
            df = pd.read_csv(uploaded_recipients)
        else:
            return tokens
        if df is None or len(df) == 0:
            return tokens
        row0 = df.iloc[0].to_dict()
        for k in ["Ten", "Email", "Code"]:
            if k in row0 and str(row0[k]).strip():
                tokens[k] = str(row0[k]).strip()
        return tokens
    except Exception:
        return tokens


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
            index=1,
        )

        host_default = "smtp.gmail.com" if smtp_provider == "Gmail (STARTTLS)" else (
            "smtp.office365.com" if smtp_provider == "Outlook (STARTTLS)" else "smtp.gmail.com"
        )

        smtp_host_default = _env_str("SMTP_HOST", "") or host_default
        smtp_port_default = _env_int("SMTP_PORT", 587)
        use_ssl_default = _env_bool("SMTP_USE_SSL", False)
        smtp_user_default = _env_str("SMTP_USER", "")
        smtp_pass_default = _env_str("SMTP_PASS", "")
        from_name_default = _env_str("FROM_NAME", "Bookmedi")
        default_subject_default = _env_str(
            "DEFAULT_SUBJECT",
            "Kết quả bài thi Versant Professional English Test  - {{Ten}}",
        )
        dry_run_default = _env_bool("DRY_RUN_DEFAULT", True)
        rate_delay_default = _env_float("RATE_DELAY_DEFAULT", 1.5)

        smtp_host = st.sidebar.text_input("SMTP Host", value=smtp_host_default)
        smtp_port = st.sidebar.number_input("SMTP Port", min_value=1, max_value=65535, value=int(smtp_port_default))
        use_ssl = st.sidebar.checkbox("Use SSL (SMTPS)", value=bool(use_ssl_default))
        smtp_user = st.sidebar.text_input("SMTP User (email)", value=smtp_user_default)
        smtp_pass = st.sidebar.text_input("SMTP Password/App Password", type="password", value=smtp_pass_default)
        from_name = st.sidebar.text_input("From Name", value=from_name_default)
        default_subject = st.sidebar.text_input(
            "Default Subject",
            value=default_subject_default,
        )
        dry_run = st.sidebar.checkbox("Dry-run (không gửi thật)", value=bool(dry_run_default))
        rate_delay_default_clamped = max(0.0, min(10.0, float(rate_delay_default)))
        rate_delay = st.sidebar.slider("Delay giữa mỗi email (giây)", 0.0, 10.0, rate_delay_default_clamped, 0.5)

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
            st.caption("Bạn có thể gõ nội dung và dùng token như {{Ten}}, {{Email}}, {{Code}} ...")
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

            # Fixed header/footer controls
            st.markdown("---")
            st.subheader("Header/Footer cố định")
            base = Path(__file__).parent
            header_file = base / "header.html"
            footer_file = base / "footer.html"
            try:
                default_header = header_file.read_text(encoding="utf-8") if header_file.exists() else ""
            except Exception:
                default_header = ""
            try:
                default_footer = footer_file.read_text(encoding="utf-8") if footer_file.exists() else ""
            except Exception:
                default_footer = ""

            with st.expander("Thiết lập Header/Footer (có preview)", expanded=False):
                st.caption("Bạn có thể soạn bằng WYSIWYG hoặc dán HTML. Khi gửi thật, hệ thống tự nhúng ảnh inline (CID) cho các <img src=\"file.png\">.")

                # Default to HTML unless CKEditor is available (Quill sometimes returns None until edited)
                hf_default = "WYSIWYG" if st_ckeditor is not None else "HTML"
                hf_editor_mode = st.radio(
                    "Cách soạn Header/Footer",
                    ["WYSIWYG", "HTML"],
                    horizontal=True,
                    index=0 if hf_default == "WYSIWYG" else 1,
                    key="hf_editor_mode",
                )

                colhf = st.columns(2)
                header_html: str
                footer_html: str

                if hf_editor_mode == "WYSIWYG" and (st_ckeditor is not None or st_quill is not None):
                    with colhf[0]:
                        st.write("Header")
                        if st_ckeditor is not None:
                            header_html = st_ckeditor(default_header, key="header_wysiwyg", height=220)
                        else:
                            # Seed initial value so Quill isn't empty on first render
                            if "header_wysiwyg" not in st.session_state:
                                st.session_state["header_wysiwyg"] = default_header
                            qv = st_quill(html=True, placeholder="Soạn header...", key="header_wysiwyg")  # type: ignore[misc]
                            header_html = default_header if _looks_empty_rich_html(qv) else str(qv)
                    with colhf[1]:
                        st.write("Footer")
                        if st_ckeditor is not None:
                            footer_html = st_ckeditor(default_footer, key="footer_wysiwyg", height=220)
                        else:
                            if "footer_wysiwyg" not in st.session_state:
                                st.session_state["footer_wysiwyg"] = default_footer
                            qv = st_quill(html=True, placeholder="Soạn footer...", key="footer_wysiwyg")  # type: ignore[misc]
                            footer_html = default_footer if _looks_empty_rich_html(qv) else str(qv)
                    # Mirror into session keys used later in sending
                    st.session_state["header_html"] = header_html
                    st.session_state["footer_html"] = footer_html
                else:
                    header_html = colhf[0].text_area("Header HTML (cố định)", value=default_header, height=220, key="header_html")
                    footer_html = colhf[1].text_area("Footer HTML (cố định)", value=default_footer, height=220, key="footer_html")

                col_actions = st.columns([2, 2, 3])
                with col_actions[0]:
                    if st.button("Lưu header.html / footer.html", key="save_hf"):
                        try:
                            header_file.write_text(header_html or "", encoding="utf-8")
                            footer_file.write_text(footer_html or "", encoding="utf-8")
                            st.success("Đã lưu header.html và footer.html")
                        except Exception as exc:
                            st.error(f"Không thể lưu header/footer: {exc}")

                with col_actions[1]:
                    if st.button("Reset về file hiện có", key="reset_hf"):
                        st.session_state.pop("header_html", None)
                        st.session_state.pop("footer_html", None)
                        st.session_state.pop("header_wysiwyg", None)
                        st.session_state.pop("footer_wysiwyg", None)
                        _safe_rerun()

                with col_actions[2]:
                    st.write("")

                # Logo quick replace (logomedi.png) used by cid:bookmedi_logo
                st.markdown("**Logo mặc định (CID `bookmedi_logo`)**")
                project_base = Path(__file__).parent.resolve()
                logo_options = []
                for name in ["logo_cdimex.png", "logomedi.png"]:
                    if (project_base / name).exists():
                        logo_options.append(name)
                if not logo_options:
                    logo_options = ["logomedi.png"]

                default_logo = st.session_state.get("cid_logo_filename") or ("logomedi.png" if "logomedi.png" in logo_options else logo_options[0])
                cid_logo_filename = st.radio(
                    "Chọn logo gửi kèm (CID `bookmedi_logo`)",
                    options=logo_options,
                    index=logo_options.index(default_logo) if default_logo in logo_options else 0,
                    horizontal=True,
                    key="cid_logo_filename",
                )

                chosen_logo_path = (project_base / cid_logo_filename).resolve()
                if chosen_logo_path.exists():
                    st.image(str(chosen_logo_path), caption=f"{cid_logo_filename}", width=220)

                st.caption("Gợi ý: trong header/footer dùng `<img src=\"cid:bookmedi_logo\">` để logo luôn nhúng inline khi gửi mail.")

                st.markdown("**Ảnh khác (banner, icon...)**")
                st.caption("Upload ảnh vào `uploads/` rồi chèn: `<img src=\"tenfile.png\">` (gửi thật sẽ tự nhúng inline).")
                img_uploads = st.file_uploader("Upload ảnh", type=["png", "jpg", "jpeg", "gif", "webp"], accept_multiple_files=True, key="img_upl")
                if img_uploads:
                    if st.button("Lưu ảnh vào uploads/", key="save_imgs"):
                        saved = 0
                        for up in img_uploads:
                            try:
                                dest = (Path(__file__).parent / "uploads" / Path(up.name).name)
                                with open(dest, "wb") as f:
                                    f.write(up.getbuffer())
                                saved += 1
                            except Exception as exc:
                                st.error(f"Không thể lưu {up.name}: {exc}")
                        if saved:
                            st.success(f"Đã lưu {saved} ảnh vào thư mục uploads/")

                st.markdown("**Xem trước (preview)**")
                tokens = _load_preview_tokens(up_recipients)
                sample_body = (html_content or "").strip() or "<p>(Nội dung trống)</p>"
                composed = _compose_email_html(header_html or "", sample_body, footer_html or "")
                composed = render_template(composed, tokens)  # type: ignore[name-defined]
                composed_preview = _preview_html_with_embedded_images(composed, base / "uploads", cid_logo_filename=cid_logo_filename)
                st.components.v1.html(composed_preview, height=520, scrolling=True)

            with st.expander("Chèn token nhanh", expanded=False):
                col_t = st.columns(3)
                tokens = ["{{Ten}}", "{{Email}}", "{{Code}}", "{{NgayGui}}"]
                for i, tk in enumerate(tokens):
                    col_idx = i % len(col_t)
                    if col_t[col_idx].button(tk, key=f"ins_{tk}"):
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
                            # Gộp header + nội dung + footer thành 1 HTML doc (tránh nested <html>/<body>)
                            hdr = st.session_state.get("header_html", "") or ""
                            ftr = st.session_state.get("footer_html", "") or ""
                            full_html = _compose_email_html(hdr, content_to_use, ftr) if (hdr or ftr) else content_to_use
                            # Lưu trong uploads/ để các đường dẫn tương đối trong HTML có thể tham chiếu tới tệp trong dự án
                            try:
                                editor_tpl = upload_dir / "_editor_template.html"
                                editor_tpl.write_text(full_html, encoding="utf-8")
                                template_path = editor_tpl
                            except Exception:
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8") as tmp_html:
                                    tmp_html.write(full_html)
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
                            cid_logo_filename=str(st.session_state.get("cid_logo_filename") or "logomedi.png"),
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
