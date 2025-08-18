from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from send_mail_merge import run_merge


class MailMergeGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Mail Merge - Bookmedi")
        self.geometry("780x560")

        # Form state
        self.recipients_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.smtp_host_var = tk.StringVar(value="smtp.gmail.com")
        self.smtp_port_var = tk.StringVar(value="587")
        self.smtp_user_var = tk.StringVar()
        self.smtp_pass_var = tk.StringVar()
        self.from_name_var = tk.StringVar(value="Bookmedi")
        self.default_subject_var = tk.StringVar(value="Kết quả bài thi Versant level 1 - {{Ten}}")
        self.use_ssl_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=True)

        self._build_form()

    def _build_form(self) -> None:
        pad = {"padx": 6, "pady": 4, "sticky": "w"}

        def add_row(row_index: int, label_text: str, widget: tk.Widget) -> None:
            tk.Label(self, text=label_text).grid(row=row_index, column=0, **pad)
            widget.grid(row=row_index, column=1, columnspan=2, sticky="ew", padx=6, pady=4)

        self.grid_columnconfigure(1, weight=1)

        # Recipients
        recipients_entry = tk.Entry(self, textvariable=self.recipients_var)
        add_row(0, "Recipients (.xlsx/.csv)", recipients_entry)
        tk.Button(self, text="Browse", command=self._pick_recipients).grid(row=0, column=3, **pad)

        # Template
        template_entry = tk.Entry(self, textvariable=self.template_var)
        add_row(1, "Template (.html)", template_entry)
        tk.Button(self, text="Browse", command=self._pick_template).grid(row=1, column=3, **pad)

        # SMTP settings
        add_row(2, "SMTP Host", tk.Entry(self, textvariable=self.smtp_host_var))
        add_row(3, "SMTP Port", tk.Entry(self, textvariable=self.smtp_port_var))
        add_row(4, "SMTP User", tk.Entry(self, textvariable=self.smtp_user_var))
        add_row(5, "SMTP Pass", tk.Entry(self, show="*", textvariable=self.smtp_pass_var))
        add_row(6, "From Name", tk.Entry(self, textvariable=self.from_name_var))
        add_row(7, "Default Subject", tk.Entry(self, textvariable=self.default_subject_var))

        # Options
        tk.Checkbutton(self, text="Use SSL (SMTPS)", variable=self.use_ssl_var).grid(row=8, column=0, **pad)
        tk.Checkbutton(self, text="Dry-run (không gửi thật)", variable=self.dry_run_var).grid(row=8, column=1, **pad)

        # Actions
        tk.Button(self, text="Send", command=self._on_send).grid(row=9, column=0, padx=6, pady=8)

        # Log area
        self.log = scrolledtext.ScrolledText(self, height=16)
        self.log.grid(row=10, column=0, columnspan=4, sticky="nsew", padx=6, pady=6)
        self.grid_rowconfigure(10, weight=1)

    def _pick_recipients(self) -> None:
        path = filedialog.askopenfilename(title="Chọn recipients", filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if path:
            self.recipients_var.set(path)

    def _pick_template(self) -> None:
        path = filedialog.askopenfilename(title="Chọn template HTML", filetypes=[("HTML", "*.html")])
        if path:
            self.template_var.set(path)

    def _append_log(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def _on_send(self) -> None:
        try:
            smtp_port = int(self.smtp_port_var.get())
        except ValueError:
            messagebox.showerror("Lỗi", "SMTP Port phải là số")
            return

        kwargs = dict(
            recipients=self.recipients_var.get(),
            template=self.template_var.get(),
            smtp_host=self.smtp_host_var.get(),
            smtp_port=smtp_port,
            smtp_user=self.smtp_user_var.get(),
            smtp_pass=self.smtp_pass_var.get(),
            from_name=self.from_name_var.get(),
            default_subject=self.default_subject_var.get(),
            rate_delay=1.5,
            dry_run=self.dry_run_var.get(),
            use_ssl=self.use_ssl_var.get(),
            progress_callback=self._append_log,
        )

        def worker() -> None:
            try:
                run_merge(**kwargs)
                self._append_log("Hoàn tất.")
            except Exception as exc:
                messagebox.showerror("Lỗi", str(exc))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = MailMergeGUI()
    # Prefill defaults from project folder
    base = Path(__file__).parent
    rec = base / "recipients.xlsx"
    tpl = base / "template.html"
    if rec.exists():
        app.recipients_var.set(str(rec))
    if tpl.exists():
        app.template_var.set(str(tpl))
    app.mainloop()


