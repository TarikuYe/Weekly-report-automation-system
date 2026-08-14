"""
Weekly Report Automator — GUI
==============================
A lightweight desktop front-end for main.py's consolidation pipeline,
with an integrated Email Automation tab for SMTP settings, department
registry, deadline configuration, and on-demand email dispatch.

Run with:  python gui.py
Requires:  main.py, email_sender.py, email_checker.py in the same folder.
"""

from __future__ import annotations

import json
import logging
import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import main  # the consolidation pipeline (main.py)
import email_checker
import email_sender

# --------------------------------------------------------------------------- #
# Palette — mirrors the styling used inside the generated report itself
# --------------------------------------------------------------------------- #
INK = "#1E2430"
MUTED = "#6B7280"
ACCENT = "#2D5BFF"
ACCENT_HOVER = "#1E46D6"
ACCENT_DARK = "#1B2A5C"
SUCCESS = "#1B8A5A"
WARNING = "#B7791F"
DANGER = "#C0392B"
BG = "#F5F6FA"
CARD_BG = "#FFFFFF"
BORDER = "#E3E6EE"
CONSOLE_BG = "#1B1F2B"
CONSOLE_FG = "#D7DAE3"

FONT_FAMILY = "Segoe UI"


class QueueHandler(logging.Handler):
    """Pushes log records into a thread-safe queue; the GUI drains it on the main thread."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record: logging.LogRecord) -> None:
        self.log_queue.put(record)


class Card(tk.Frame):
    """A simple bordered white panel used to group related controls."""

    def __init__(self, parent, title: str = "", **kwargs):
        super().__init__(parent, bg=CARD_BG, highlightbackground=BORDER,
                          highlightthickness=1, bd=0, **kwargs)
        if title:
            header = tk.Label(
                self, text=title, bg=CARD_BG, fg=ACCENT_DARK,
                font=(FONT_FAMILY, 10, "bold"), anchor="w",
            )
            header.pack(fill="x", padx=16, pady=(14, 4))
        self.body = tk.Frame(self, bg=CARD_BG)
        self.body.pack(fill="both", expand=True, padx=16, pady=(0, 14))


class WeeklyReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Weekly Report Automator")
        self.root.geometry("800x700")
        self.root.minsize(720, 600)
        self.root.configure(bg=BG)

        self.config_path = main.base_dir() / "config.json"
        self.config = self._load_config()

        self.log_queue: queue.Queue = queue.Queue()
        self.is_running = False
        self._alive = True          # set to False once the window is closing
        self.last_output_path: Path | None = None

        self._build_style()
        self._build_layout()
        self._wire_logging()
        self._poll_log_queue()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Start background email scheduler (runs even while GUI is open)
        email_checker.schedule_auto_check(self.config_path)
        self._poll_scheduler_status()

    # ------------------------------------------------------------------ #
    # Window lifecycle
    # ------------------------------------------------------------------ #
    def _on_close(self) -> None:
        if self.is_running:
            if not messagebox.askyesno(
                "Pipeline running",
                "The pipeline is still running in the background.\n"
                "Close anyway? The output file may be incomplete.",
            ):
                return
        email_checker.stop_auto_check()
        self._alive = False
        self.root.destroy()

    # ------------------------------------------------------------------ #
    # Config
    # ------------------------------------------------------------------ #
    def _load_config(self) -> dict:
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            base = main.base_dir()
            return {
                "input_dir": str(base / "input"),
                "output_file": str(base / "output" / "master_weekly_report.xlsx"),
            }

    def _save_config(self) -> None:
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=4)

    # ------------------------------------------------------------------ #
    # Style
    # ------------------------------------------------------------------ #
    def _build_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Modern.Horizontal.TProgressbar",
                         troughcolor=BORDER, background=ACCENT,
                         bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)

        style.configure("TScrollbar", background=BORDER, troughcolor=CARD_BG,
                         bordercolor=CARD_BG, arrowcolor=MUTED)

        style.configure("Tab.TNotebook", background=BG, borderwidth=0)
        style.configure("Tab.TNotebook.Tab",
                         background=CARD_BG, foreground=MUTED,
                         padding=[16, 8], font=(FONT_FAMILY, 9, "bold"))
        style.map("Tab.TNotebook.Tab",
                  background=[("selected", ACCENT_DARK)],
                  foreground=[("selected", "white")])

    # ------------------------------------------------------------------ #
    # Layout
    # ------------------------------------------------------------------ #
    def _build_layout(self) -> None:
        self._build_header()

        # Tab container
        self._notebook = ttk.Notebook(self.root, style="Tab.TNotebook")
        self._notebook.pack(fill="both", expand=True, padx=0, pady=0)

        # ---- Tab 1: Pipeline ----
        pipeline_tab = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(pipeline_tab, text="  ⚡ Pipeline  ")

        body = tk.Frame(pipeline_tab, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)
        self._build_paths_card(body)
        self._build_action_row(body)
        self._build_console_card(body)

        # ---- Tab 2: Email Automation ----
        email_tab = tk.Frame(self._notebook, bg=BG)
        self._notebook.add(email_tab, text="  ✉ Email Automation  ")
        self._build_email_tab(email_tab)

        self._build_status_bar()

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=ACCENT_DARK, height=84)
        header.pack(fill="x")
        header.pack_propagate(False)

        text_col = tk.Frame(header, bg=ACCENT_DARK)
        text_col.pack(side="left", padx=24, pady=14)

        tk.Label(
            text_col, text="⚡ Weekly Report Automator", fg="white", bg=ACCENT_DARK,
            font=(FONT_FAMILY, 16, "bold"),
        ).pack(anchor="w")
        tk.Label(
            text_col, text="Merge every department file into one polished master report",
            fg="#B9C2E0", bg=ACCENT_DARK, font=(FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(2, 0))

    def _build_paths_card(self, parent: tk.Frame) -> None:
        card = Card(parent, title="FILE PATH SETTINGS")
        card.pack(fill="x", pady=(0, 14))

        grid = card.body
        grid.columnconfigure(1, weight=1)

        # Input folder
        tk.Label(grid, text="Input Folder", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, sticky="w", pady=(4, 2))
        self.input_entry = self._make_entry(grid)
        self.input_entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        self.input_entry.insert(0, self.config.get("input_dir", ""))
        self._make_browse_button(grid, self.browse_input).grid(row=1, column=2)

        self.file_count_label = tk.Label(
            grid, text="", bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 8, "italic"),
        )
        self.file_count_label.grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 10))

        # Output file
        tk.Label(grid, text="Save Output As", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 9, "bold")).grid(row=3, column=0, sticky="w", pady=(4, 2))
        self.output_entry = self._make_entry(grid)
        self.output_entry.grid(row=4, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        self.output_entry.insert(0, self.config.get("output_file", ""))
        self._make_browse_button(grid, self.browse_output).grid(row=4, column=2)

        self._refresh_file_count()

    def _make_entry(self, parent) -> tk.Entry:
        return tk.Entry(
            parent, font=(FONT_FAMILY, 10), bd=1, relief="solid",
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )

    def _make_browse_button(self, parent, command) -> tk.Button:
        return tk.Button(
            parent, text="Browse…", font=(FONT_FAMILY, 9), command=command,
            bg="#EDEFF6", fg=INK, activebackground=BORDER, relief="flat",
            padx=12, pady=4, cursor="hand2",
        )

    def _build_action_row(self, parent: tk.Frame) -> None:
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 14))

        self.run_btn = tk.Button(
            row, text="Run Merger Pipeline", font=(FONT_FAMILY, 11, "bold"),
            bg=ACCENT, fg="white", activebackground=ACCENT_HOVER, activeforeground="white",
            relief="flat", height=2, cursor="hand2", command=self.start_pipeline_thread,
        )
        self.run_btn.pack(side="left", fill="x", expand=True)
        self.run_btn.bind("<Enter>", lambda e: self.run_btn.configure(bg=ACCENT_HOVER))
        self.run_btn.bind("<Leave>", lambda e: self.run_btn.configure(
            bg=ACCENT if not self.is_running else ACCENT_HOVER))

        self.open_output_btn = tk.Button(
            row, text="Open Output Folder", font=(FONT_FAMILY, 10), command=self.open_output_folder,
            bg="#EDEFF6", fg=INK, relief="flat", height=2, padx=14, state="disabled", cursor="hand2",
        )
        self.open_output_btn.pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(
            parent, style="Modern.Horizontal.TProgressbar", mode="indeterminate",
        )
        # gridded/packed dynamically only while running

    def _build_console_card(self, parent: tk.Frame) -> None:
        card = Card(parent, title="EXECUTION LOG")
        card.pack(fill="both", expand=True)
        self._console_card = card

        console_frame = tk.Frame(card.body, bg=CONSOLE_BG)
        console_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            console_frame, font=("Consolas", 9), bg=CONSOLE_BG, fg=CONSOLE_FG,
            state="disabled", wrap="word", bd=0, padx=10, pady=8,
        )
        self.log_text.pack(fill="both", expand=True, side="left")

        self.log_text.tag_configure("INFO", foreground="#8FB4FF")
        self.log_text.tag_configure("WARNING", foreground="#F2C572")
        self.log_text.tag_configure("ERROR", foreground="#FF8080")
        self.log_text.tag_configure("timestamp", foreground="#6B7280")

        scrollbar = ttk.Scrollbar(console_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(fill="y", side="right")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        clear_btn = tk.Button(
            card.body, text="Clear Log", font=(FONT_FAMILY, 8), command=self.clear_log,
            bg=CARD_BG, fg=MUTED, relief="flat", cursor="hand2",
        )
        clear_btn.pack(anchor="e", pady=(6, 0))

    def _build_status_bar(self) -> None:
        bar = tk.Frame(self.root, bg="#EDEFF6", height=30)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self.status_dot = tk.Label(bar, text="●", fg=MUTED, bg="#EDEFF6", font=(FONT_FAMILY, 10))
        self.status_dot.pack(side="left", padx=(14, 4))

        self.status_label = tk.Label(
            bar, text="Ready", fg=MUTED, bg="#EDEFF6", font=(FONT_FAMILY, 9),
        )
        self.status_label.pack(side="left")

    # ------------------------------------------------------------------ #
    # Browsing
    # ------------------------------------------------------------------ #
    def browse_input(self) -> None:
        folder = filedialog.askdirectory(initialdir=str(main.base_dir()))
        if folder:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, folder)
            self.config["input_dir"] = folder
            self._save_config()
            self._refresh_file_count()

    def browse_output(self) -> None:
        default_dir = main.base_dir() / "output"
        default_dir.mkdir(exist_ok=True)
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")],
            initialdir=str(default_dir),
        )
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)
            self.config["output_file"] = file_path
            self._save_config()

    def _refresh_file_count(self) -> None:
        folder = Path(self.input_entry.get())
        try:
            count = 0

            # Root-level files
            if folder.exists():
                count += len([p for p in folder.glob("*.xlsx") if not p.name.startswith("~$")])

            # Per-department folders (may be external absolute paths)
            for dept in self.config.get("departments", []):
                folder_value = dept.get("folder", "").strip()
                if not folder_value:
                    continue
                dept_path = Path(folder_value) if Path(folder_value).is_absolute() else folder / folder_value
                if not dept_path.is_dir() or dept_path == folder:
                    continue
                count += len([
                    p for p in dept_path.rglob("*.xlsx")
                    if not p.name.startswith("~$")
                    and "archive" not in [part.lower() for part in p.relative_to(dept_path).parts[:-1]]
                ])

            self.file_count_label.configure(
                text=f"{count} .xlsx file(s) detected" if count else "No .xlsx files found in this folder yet"
            )
        except Exception:
            self.file_count_label.configure(text="")

    # ------------------------------------------------------------------ #
    # Logging plumbing
    # ------------------------------------------------------------------ #
    def _wire_logging(self) -> None:
        handler = QueueHandler(self.log_queue)
        handler.setFormatter(logging.Formatter("%(asctime)s|%(levelname)s|%(message)s", "%H:%M:%S"))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def _poll_log_queue(self) -> None:
        if not self._alive:
            return
        try:
            while True:
                record = self.log_queue.get_nowait()
                self._append_log(record)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def _append_log(self, record: logging.LogRecord) -> None:
        ts, level, msg = self._format_record(record)
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"{ts}  ", "timestamp")
        self.log_text.insert(tk.END, f"{msg}\n", level if level in ("INFO", "WARNING", "ERROR") else "")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)

    @staticmethod
    def _format_record(record: logging.LogRecord) -> tuple[str, str, str]:
        formatted = f"{record.levelname}|{record.getMessage()}"
        level, _, msg = formatted.partition("|")
        ts = time.strftime("%H:%M:%S")
        return ts, level, msg

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------------ #
    # Pipeline execution
    # ------------------------------------------------------------------ #
    def start_pipeline_thread(self) -> None:
        if self.is_running:
            return

        input_dir = self.input_entry.get().strip()
        output_file = self.output_entry.get().strip()

        if not input_dir or not output_file:
            messagebox.showwarning("Missing Paths", "Please set both an input folder and an output file.")
            return

        self.config["input_dir"] = input_dir
        self.config["output_file"] = output_file
        self._save_config()

        self.is_running = True
        self.run_btn.configure(state="disabled", text="Processing… please wait", bg=ACCENT_HOVER)
        self.open_output_btn.configure(state="disabled")
        self._set_status("Running pipeline…", WARNING)
        self.progress.pack(fill="x", pady=(0, 10), before=self._console_card)
        self.progress.start(12)

        threading.Thread(target=self._run_pipeline_worker, args=(input_dir, output_file), daemon=True).start()

    def _run_pipeline_worker(self, input_dir: str, output_file: str) -> None:
        try:
            departments = self.config.get("departments", None)
            main.run_pipeline_from_paths(input_dir, output_file, departments)
            self.last_output_path = Path(output_file)
            # Find the most-recently-created archive batch for the confirmation message.
            # With per-department archiving, check both the root input/archive/ and
            # each department folder's own archive/ subdirectory.
            archive_batch: Path | None = None
            newest_mtime: float = 0.0

            def _check_archive_dir(archive_dir: Path) -> None:
                nonlocal archive_batch, newest_mtime
                if not archive_dir.exists():
                    return
                for d in archive_dir.iterdir():
                    if d.is_dir():
                        mtime = d.stat().st_mtime
                        if mtime > newest_mtime:
                            newest_mtime = mtime
                            archive_batch = d

            # Root-level archive
            _check_archive_dir(Path(input_dir) / "archive")

            # Per-department archives
            for dept in (departments or []):
                folder_value = dept.get("folder", "").strip()
                if not folder_value:
                    continue
                dept_path = (Path(folder_value) if Path(folder_value).is_absolute()
                             else Path(input_dir) / folder_value)
                if dept_path.is_dir():
                    _check_archive_dir(dept_path / "archive")
            if self._alive:
                self.root.after(0, lambda: self._on_pipeline_success(archive_batch))
        except Exception as exc:
            if self._alive:
                self.root.after(0, lambda: self._on_pipeline_error(exc))

    def _on_pipeline_success(self, archive_batch: Path | None = None) -> None:
        self.is_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal", text="Run Merger Pipeline", bg=ACCENT)
        self.open_output_btn.configure(state="normal")
        self._set_status("Completed successfully", SUCCESS)
        self._refresh_file_count()

        archive_note = (
            f"\n\nInput files archived to:\n{archive_batch}"
            if archive_batch
            else "\n\n(No input files were found to archive.)"
        )
        messagebox.showinfo(
            "Pipeline Complete",
            "All weekly files were merged and the modern report styling has been applied."
            + archive_note,
        )

    def _on_pipeline_error(self, exc: Exception) -> None:
        self.is_running = False
        self.progress.stop()
        self.progress.pack_forget()
        self.run_btn.configure(state="normal", text="Run Merger Pipeline", bg=ACCENT)
        self._set_status("Pipeline failed", DANGER)
        messagebox.showerror("Error", f"Failed to execute pipeline:\n{exc}")

    def _set_status(self, text: str, color: str) -> None:
        self.status_dot.configure(fg=color)
        self.status_label.configure(text=text, fg=INK if color != MUTED else MUTED)

    def open_output_folder(self) -> None:
        if not self.last_output_path:
            return
        folder = self.last_output_path.parent.resolve()
        try:
            import os
            import platform
            import subprocess

            system = platform.system()
            if system == "Windows":
                os.startfile(folder)  # type: ignore[attr-defined]
            elif system == "Darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:
            messagebox.showwarning("Couldn't open folder", str(exc))

    # ================================================================== #
    #  EMAIL AUTOMATION TAB
    # ================================================================== #

    def _build_email_tab(self, parent: tk.Frame) -> None:
        """Scrollable Email Automation tab."""
        canvas = tk.Canvas(parent, bg=BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self._email_scroll_frame = tk.Frame(canvas, bg=BG)

        self._email_scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self._email_scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        body = self._email_scroll_frame
        pad  = {"fill": "x", "padx": 20}

        self._build_email_toggle_card(body).pack(**pad, pady=(16, 14))
        self._build_smtp_card(body).pack(**pad, pady=(0, 14))
        self._build_deadline_card(body).pack(**pad, pady=(0, 14))
        self._build_departments_card(body).pack(**pad, pady=(0, 14))
        self._build_auto_send_card(body).pack(**pad, pady=(0, 14))
        self._build_email_actions_card(body).pack(**pad, pady=(0, 14))
        self._build_submission_status_card(body).pack(**pad, pady=(0, 20))


    # ------------------------------------------------------------------ #
    # Email Tab — Toggle card
    # ------------------------------------------------------------------ #
    def _build_email_toggle_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="EMAIL AUTOMATION")
        g = card.body
        g.columnconfigure(1, weight=1)

        self._email_enabled_var = tk.BooleanVar(
            value=self.config.get("email", {}).get("enabled", False)
        )
        tk.Label(g, text="Enable email automation", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 10)).grid(row=0, column=0, sticky="w", pady=4)
        toggle = ttk.Checkbutton(
            g, variable=self._email_enabled_var,
            command=self._on_email_toggle,
        )
        toggle.grid(row=0, column=1, sticky="w", padx=8)

        tk.Label(
            g,
            text="When enabled, the system checks department folders and sends reminders\n"
                 "to departments that missed the deadline, and acknowledgments to those who submitted.",
            bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 9), justify="left",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 4))
        return card

    def _on_email_toggle(self) -> None:
        if "email" not in self.config:
            self.config["email"] = {}
        self.config["email"]["enabled"] = self._email_enabled_var.get()
        self._save_config()

    # ------------------------------------------------------------------ #
    # Email Tab — SMTP card
    # ------------------------------------------------------------------ #
    def _build_smtp_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="SMTP SETTINGS")
        g = card.body
        g.columnconfigure(1, weight=1)
        g.columnconfigure(3, weight=1)

        ecfg = self.config.get("email", {})

        def row(label, var, r, c=0, show=""):
            tk.Label(g, text=label, bg=CARD_BG, fg=INK,
                     font=(FONT_FAMILY, 9, "bold")).grid(
                row=r, column=c, sticky="w", pady=(6, 2), padx=(0, 8))
            e = tk.Entry(g, textvariable=var, font=(FONT_FAMILY, 10),
                         bd=1, relief="solid", show=show,
                         highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=ACCENT)
            e.grid(row=r + 1, column=c, sticky="ew", padx=(0, 12))
            return e

        self._smtp_host  = tk.StringVar(value=ecfg.get("smtp_host",  "smtp.gmail.com"))
        self._smtp_port  = tk.StringVar(value=str(ecfg.get("smtp_port", 587)))
        self._smtp_user  = tk.StringVar(value=ecfg.get("smtp_user",  ""))
        self._smtp_pass  = tk.StringVar(value=ecfg.get("smtp_password", ""))
        self._smtp_from  = tk.StringVar(value=ecfg.get("sender_email", ""))
        self._smtp_name  = tk.StringVar(value=ecfg.get("sender_name",  "Weekly Report System"))
        self._smtp_tls   = tk.BooleanVar(value=bool(ecfg.get("use_tls", True)))

        row("SMTP Host",          self._smtp_host, 0, 0)
        row("Port",               self._smtp_port, 0, 2)
        row("Login / Username",   self._smtp_user, 2, 0)
        row("Password / App Key", self._smtp_pass, 2, 2, show="•")
        row("Sender Email",       self._smtp_from, 4, 0)
        row("Sender Display Name",self._smtp_name, 4, 2)

        tls_row = tk.Frame(g, bg=CARD_BG)
        tls_row.grid(row=6, column=0, columnspan=4, sticky="w", pady=(10, 2))
        ttk.Checkbutton(tls_row, variable=self._smtp_tls,
                        text=" Use STARTTLS (port 587) — uncheck for SSL port 465").pack(side="left")

        btn_row = tk.Frame(g, bg=CARD_BG)
        btn_row.grid(row=7, column=0, columnspan=4, sticky="w", pady=(12, 4))

        tk.Button(btn_row, text="Save SMTP Settings",
                  font=(FONT_FAMILY, 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  padx=14, pady=5, cursor="hand2",
                  command=self._save_smtp).pack(side="left")

        tk.Button(btn_row, text="Send Test Email…",
                  font=(FONT_FAMILY, 9),
                  bg="#EDEFF6", fg=INK, relief="flat",
                  padx=14, pady=5, cursor="hand2",
                  command=self._send_test_email).pack(side="left", padx=(10, 0))

        self._smtp_status = tk.Label(btn_row, text="", bg=CARD_BG, fg=MUTED,
                                     font=(FONT_FAMILY, 9, "italic"))
        self._smtp_status.pack(side="left", padx=(12, 0))
        return card

    def _save_smtp(self) -> None:
        if "email" not in self.config:
            self.config["email"] = {}
        e = self.config["email"]
        e["smtp_host"]     = self._smtp_host.get().strip()
        e["smtp_port"]     = int(self._smtp_port.get().strip() or 587)
        e["smtp_user"]     = self._smtp_user.get().strip()
        e["smtp_password"] = self._smtp_pass.get()
        e["sender_email"]  = self._smtp_from.get().strip()
        e["sender_name"]   = self._smtp_name.get().strip()
        e["use_tls"]       = self._smtp_tls.get()
        self._save_config()
        self._smtp_status.configure(text="✓ Saved", fg=SUCCESS)
        self.root.after(3000, lambda: self._smtp_status.configure(text=""))

    def _send_test_email(self) -> None:
        self._save_smtp()
        to = self._smtp_from.get().strip() or self._smtp_user.get().strip()
        if not to:
            messagebox.showwarning("No Address",
                                   "Enter a Sender Email first — the test email will be sent to it.")
            return
        self._smtp_status.configure(text="Sending…", fg=WARNING)
        self.root.update_idletasks()

        def _worker():
            ok, msg = email_sender.send_test_email(self.config, to)
            def _ui():
                if ok:
                    self._smtp_status.configure(text="✓ Test email sent!", fg=SUCCESS)
                    messagebox.showinfo("Test Sent", f"Test email delivered to:\n{to}")
                else:
                    self._smtp_status.configure(text="✗ Failed", fg=DANGER)
                    messagebox.showerror("Send Failed", msg)
            self.root.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    # ------------------------------------------------------------------ #
    # Email Tab — Deadline card
    # ------------------------------------------------------------------ #
    def _build_deadline_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="SUBMISSION DEADLINE")
        g = card.body
        g.columnconfigure(3, weight=1)

        ecfg = self.config.get("email", {})
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

        # ── Deadline Day ──
        tk.Label(g, text="Deadline Day", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=0, sticky="w", pady=(4, 2))
        self._deadline_day = tk.StringVar(value=ecfg.get("deadline_day", "Thursday"))
        day_cb = ttk.Combobox(g, textvariable=self._deadline_day,
                              values=days, state="readonly", width=14,
                              font=(FONT_FAMILY, 10))
        day_cb.grid(row=1, column=0, sticky="w", padx=(0, 16))

        # ── Deadline Hour ──
        tk.Label(g, text="Hour (24h)", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=1, sticky="w", pady=(4, 2))
        self._deadline_hour = tk.StringVar(value=str(ecfg.get("deadline_hour", 17)))
        hour_cb = ttk.Combobox(g, textvariable=self._deadline_hour,
                               values=[str(h) for h in range(0, 24)],
                               state="readonly", width=6,
                               font=(FONT_FAMILY, 10))
        hour_cb.grid(row=1, column=1, sticky="w", padx=(0, 8))

        # ── Separator colon ──
        tk.Label(g, text=":", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 12, "bold")).grid(row=1, column=2, sticky="w", padx=(0, 8))

        # ── Deadline Minute ──
        tk.Label(g, text="Minute", bg=CARD_BG, fg=INK,
                 font=(FONT_FAMILY, 9, "bold")).grid(row=0, column=3, sticky="w", pady=(4, 2))
        self._deadline_minute = tk.StringVar(value=str(ecfg.get("deadline_minute", 0)))
        min_cb = ttk.Combobox(g, textvariable=self._deadline_minute,
                              values=[f"{m:02d}" for m in range(0, 60, 5)] + ["59"],
                              state="readonly", width=6,
                              font=(FONT_FAMILY, 10))
        min_cb.grid(row=1, column=3, sticky="w", padx=(0, 16))

        # ── Info label ──
        self._deadline_info = tk.Label(g, text="", bg=CARD_BG, fg=MUTED,
                                       font=(FONT_FAMILY, 9, "italic"))
        self._deadline_info.grid(row=2, column=0, columnspan=5, sticky="w", pady=(8, 2))
        self._refresh_deadline_info()

        # Auto-save on any selection change
        for cb in (day_cb, hour_cb, min_cb):
            cb.bind("<<ComboboxSelected>>", lambda e: self._save_deadline())

        tk.Button(g, text="Save Deadline",
                  font=(FONT_FAMILY, 9, "bold"),
                  bg=ACCENT, fg="white", relief="flat",
                  padx=14, pady=5, cursor="hand2",
                  command=self._save_deadline).grid(row=3, column=0,
                                                    columnspan=2, sticky="w",
                                                    pady=(10, 4))
        return card

    def _save_deadline(self) -> None:
        if "email" not in self.config:
            self.config["email"] = {}
        self.config["email"]["deadline_day"]    = self._deadline_day.get()
        self.config["email"]["deadline_hour"]   = int(self._deadline_hour.get())
        self.config["email"]["deadline_minute"] = int(self._deadline_minute.get())
        self._save_config()
        self._refresh_deadline_info()

    def _refresh_deadline_info(self) -> None:
        try:
            nd   = email_checker.next_deadline(self.config)
            past = email_checker.is_past_deadline(self.config)
            status = "⚠ Past deadline this week" if past else "Next deadline"
            color  = WARNING if past else SUCCESS
            self._deadline_info.configure(
                text=f"{status}: {nd.strftime('%A %d %b %Y at %H:%M')}",
                fg=color,
            )
        except Exception:
            pass

    # ------------------------------------------------------------------ #
    # Email Tab — Departments & Employees card
    # ------------------------------------------------------------------ #
    def _build_departments_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="DEPARTMENTS & EMPLOYEES")
        g = card.body
        g.columnconfigure(0, weight=1)

        # Instruction line
        tk.Label(
            g,
            text="Each department has its own folder. Add every employee who must submit "
                 "a weekly report.\n"
                 "Set 'Expected File' to the exact filename for that person, "
                 "or leave blank to accept any .xlsx in that folder.",
            bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 8), justify="left",
        ).grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Scrollable inner frame for departments
        self._dept_outer = tk.Frame(g, bg=BG, bd=1, relief="solid")
        self._dept_outer.grid(row=1, column=0, sticky="ew", pady=(0, 6))

        self._dept_blocks: list[dict] = []   # one entry per department
        self._render_dept_blocks()

        btn_row = tk.Frame(g, bg=CARD_BG)
        btn_row.grid(row=2, column=0, sticky="w", pady=(8, 4))

        tk.Button(
            btn_row, text="+ Add Department",
            font=(FONT_FAMILY, 9), bg="#EDEFF6", fg=INK,
            relief="flat", padx=12, pady=4, cursor="hand2",
            command=self._add_dept_block,
        ).pack(side="left")

        tk.Button(
            btn_row, text="Save All",
            font=(FONT_FAMILY, 9, "bold"),
            bg=ACCENT, fg="white", relief="flat",
            padx=14, pady=5, cursor="hand2",
            command=self._save_departments,
        ).pack(side="left", padx=(10, 0))

        self._dept_save_status = tk.Label(
            btn_row, text="", bg=CARD_BG, fg=MUTED,
            font=(FONT_FAMILY, 9, "italic"),
        )
        self._dept_save_status.pack(side="left", padx=(10, 0))
        return card

    # ── Render all department blocks from self.config ──────────────────
    def _render_dept_blocks(self) -> None:
        """Destroy and recreate all department blocks from config."""
        for w in self._dept_outer.winfo_children():
            w.destroy()
        self._dept_blocks.clear()
        for dept in self.config.get("departments", []):
            self._add_dept_block(dept)

    # ── Single department block ────────────────────────────────────────
    def _add_dept_block(self, dept: dict | None = None) -> None:
        dept = dept or {}
        outer = self._dept_outer

        # ── Dept header frame ──
        blk = tk.Frame(outer, bg=CARD_BG, bd=0,
                       highlightbackground=BORDER, highlightthickness=1)
        blk.pack(fill="x", padx=4, pady=4)
        blk.columnconfigure(1, weight=1)
        blk.columnconfigure(3, weight=1)

        # Header band
        hdr = tk.Frame(blk, bg="#EDEFF6")
        hdr.pack(fill="x")

        dept_name_var       = tk.StringVar(value=dept.get("name",   ""))
        dept_folder_var     = tk.StringVar(value=dept.get("folder", ""))
        dept_cc_var         = tk.StringVar(value=dept.get("cc_email", ""))
        dept_head_email_var = tk.StringVar(value=dept.get("dept_head_email", ""))
        dept_head_pass_var  = tk.StringVar(value=dept.get("dept_head_password", ""))

        def _lbl(parent, text):
            return tk.Label(parent, text=text, bg="#EDEFF6", fg=MUTED,
                            font=(FONT_FAMILY, 8, "bold"))

        def _ent(parent, var, w, show=""):
            return tk.Entry(parent, textvariable=var, font=(FONT_FAMILY, 9),
                            bd=1, relief="solid", width=w,
                            highlightthickness=1,
                            highlightbackground=BORDER, highlightcolor=ACCENT,
                            bg="#EDEFF6", show=show)

        # Row 0: column headers — name / folder / cc_email
        _lbl(hdr, "Department Name").grid(row=0, column=0, sticky="w",
                                          padx=(8, 4), pady=(6, 0))
        _lbl(hdr, "Input Folder (path)").grid(row=0, column=1, columnspan=2, sticky="w",
                                       padx=(0, 4), pady=(6, 0))
        _lbl(hdr, "Dept-Head CC Email (optional)").grid(row=0, column=3,
                                                         sticky="w",
                                                         padx=(0, 8), pady=(6, 0))
        # Row 1: entries + browse button
        _ent(hdr, dept_name_var,   16).grid(row=1, column=0, sticky="ew",
                                            padx=(8, 4), pady=(2, 4))

        folder_entry = _ent(hdr, dept_folder_var, 20)
        folder_entry.grid(row=1, column=1, sticky="ew", padx=(0, 2), pady=(2, 4))

        def _browse_dept_folder(_var=dept_folder_var) -> None:
            initial = _var.get().strip() or str(main.base_dir())
            # Use an existing path as starting point if possible
            from pathlib import Path as _Path
            if not _Path(initial).exists():
                initial = str(main.base_dir())
            chosen = filedialog.askdirectory(
                title="Select Input Folder for this Department",
                initialdir=initial,
            )
            if chosen:
                _var.set(chosen)

        tk.Button(
            hdr, text="Browse…", font=(FONT_FAMILY, 8),
            bg="#EDEFF6", fg=INK, activebackground=BORDER,
            relief="flat", padx=8, pady=3, cursor="hand2",
            command=_browse_dept_folder,
        ).grid(row=1, column=2, padx=(0, 4), pady=(2, 4), sticky="w")

        _ent(hdr, dept_cc_var, 24).grid(row=1, column=3, sticky="ew",
                                        padx=(0, 0), pady=(2, 4))

        # Remove-dept button
        def _remove_dept(b=blk, dv=dept_name_var):
            b.destroy()
            self._dept_blocks = [x for x in self._dept_blocks
                                 if x["block"] is not b]

        tk.Button(hdr, text="✕ Remove Dept",
                  font=(FONT_FAMILY, 8), bg="#EDEFF6", fg=DANGER,
                  relief="flat", cursor="hand2",
                  command=_remove_dept).grid(row=1, column=4,
                                             padx=(8, 8), pady=(2, 4))

        # ── Row 2: Dept-Head Access (email + password for master Excel) ──
        _lbl(hdr, "🔒 Dept-Head Email (master report recipient)").grid(
            row=2, column=0, columnspan=2, sticky="w", padx=(8, 4), pady=(2, 0))
        _lbl(hdr, "🔑 Master Report Password (leave blank = no encryption)").grid(
            row=2, column=2, columnspan=3, sticky="w", padx=(0, 8), pady=(2, 0))

        _ent(hdr, dept_head_email_var, 28).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(8, 4), pady=(2, 8))

        pass_entry = _ent(hdr, dept_head_pass_var, 24, show="•")
        pass_entry.grid(row=3, column=2, columnspan=2, sticky="ew",
                        padx=(0, 4), pady=(2, 8))

        # Toggle password visibility
        _pass_visible = [False]
        def _toggle_pass_visibility(e=pass_entry, v=_pass_visible):
            v[0] = not v[0]
            e.config(show="" if v[0] else "•")

        tk.Button(
            hdr, text="👁", font=(FONT_FAMILY, 9),
            bg="#EDEFF6", fg=MUTED, relief="flat", cursor="hand2",
            command=_toggle_pass_visibility,
        ).grid(row=3, column=4, padx=(0, 8), pady=(2, 8), sticky="w")

        # ── Employee sub-frame ──
        emp_frame = tk.Frame(blk, bg=CARD_BG)
        emp_frame.pack(fill="x", padx=8, pady=(4, 0))

        # Shared grid container — headers and data rows all go here
        # so every column is sized identically.
        emp_grid = tk.Frame(emp_frame, bg=CARD_BG)
        emp_grid.pack(fill="x")
        emp_grid.columnconfigure(0, weight=3, uniform="emp")   # Name
        emp_grid.columnconfigure(1, weight=4, uniform="emp")   # Email
        emp_grid.columnconfigure(2, weight=4, uniform="emp")   # Expected File
        emp_grid.columnconfigure(3, weight=0)                  # ✕ button

        # Employee column headers (row 0 of emp_grid)
        for col_idx, col_txt in enumerate([
            "  Employee Name",
            "Email Address",
            "Expected File (leave blank = any)",
        ]):
            tk.Label(emp_grid, text=col_txt, bg=CARD_BG, fg=MUTED,
                     font=(FONT_FAMILY, 8, "bold"),
                     anchor="w").grid(row=0, column=col_idx, sticky="ew",
                                      padx=(0, 4), pady=(2, 2))

        emp_rows: list[dict] = []
        # Track the next available grid row (row 0 = headers)
        _next_row = [1]

        def _add_emp_row(emp: dict | None = None) -> None:
            emp = emp or {}
            rv  = {
                "name":  tk.StringVar(value=emp.get("name",  "")),
                "email": tk.StringVar(value=emp.get("email", "")),
                "file":  tk.StringVar(value=emp.get("expected_file", "")),
            }

            def _e(parent, var):
                return tk.Entry(
                    parent, textvariable=var,
                    font=(FONT_FAMILY, 9), bd=1, relief="solid",
                    highlightthickness=1,
                    highlightbackground=BORDER, highlightcolor=ACCENT,
                )

            r = _next_row[0]
            _next_row[0] += 1

            name_e  = _e(emp_grid, rv["name"])
            email_e = _e(emp_grid, rv["email"])
            file_e  = _e(emp_grid, rv["file"])

            name_e .grid(row=r, column=0, sticky="ew", padx=(0, 4), pady=2)
            email_e.grid(row=r, column=1, sticky="ew", padx=(0, 4), pady=2)
            file_e .grid(row=r, column=2, sticky="ew", padx=(0, 4), pady=2)

            def _rem_emp(ne=name_e, ee=email_e, fe=file_e, r=rv):
                ne.destroy(); ee.destroy(); fe.destroy()
                btn_ref[0].destroy()
                emp_rows.remove(r)

            btn = tk.Button(emp_grid, text="✕",
                            font=(FONT_FAMILY, 8), bg=CARD_BG, fg=DANGER,
                            relief="flat", cursor="hand2",
                            command=_rem_emp)
            btn.grid(row=r, column=3, padx=(2, 0), pady=2)
            btn_ref = [btn]

            emp_rows.append(rv)

        # Load existing employees
        for emp in dept.get("employees", []):
            _add_emp_row(emp)

        # "+ Add Employee" button
        add_emp_frame = tk.Frame(emp_frame, bg=CARD_BG)
        add_emp_frame.pack(fill="x", pady=(4, 8))
        tk.Button(
            add_emp_frame, text="+ Add Employee",
            font=(FONT_FAMILY, 8), bg="#F0F4FF", fg=ACCENT,
            relief="flat", padx=8, pady=3, cursor="hand2",
            command=_add_emp_row,
        ).pack(side="left")

        # Store block state
        self._dept_blocks.append({
            "block":            blk,
            "name":             dept_name_var,
            "folder":           dept_folder_var,
            "cc_email":         dept_cc_var,
            "dept_head_email":  dept_head_email_var,
            "dept_head_password": dept_head_pass_var,
            "emp_rows":         emp_rows,
        })

    # ── Collect & save ─────────────────────────────────────────────────
    def _save_departments(self) -> None:
        depts = []
        for blk in self._dept_blocks:
            name = blk["name"].get().strip()
            if not name:
                continue
            employees = []
            for rv in blk["emp_rows"]:
                emp_name = rv["name"].get().strip()
                if not emp_name:
                    continue
                employees.append({
                    "name":          emp_name,
                    "email":         rv["email"].get().strip(),
                    "expected_file": rv["file"].get().strip(),
                })
            depts.append({
                "name":               name,
                "folder":             blk["folder"].get().strip(),
                "cc_email":           blk["cc_email"].get().strip(),
                "dept_head_email":    blk["dept_head_email"].get().strip(),
                "dept_head_password": blk["dept_head_password"].get(),
                "employees":          employees,
            })
        self.config["departments"] = depts
        self._save_config()
        self._dept_save_status.configure(text="✓ Saved", fg=SUCCESS)
        self.root.after(3000, lambda: self._dept_save_status.configure(text=""))

    # ------------------------------------------------------------------ #
    # Email Tab — Auto-Send card
    # ------------------------------------------------------------------ #
    def _build_auto_send_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="AUTO-SEND SCHEDULER")
        g = card.body

        ac = self.config.get("auto_check", {})

        # ── Enable toggle row ──
        top = tk.Frame(g, bg=CARD_BG)
        top.pack(fill="x", pady=(0, 10))

        self._auto_enabled_var = tk.BooleanVar(value=ac.get("enabled", False))
        ttk.Checkbutton(
            top, variable=self._auto_enabled_var,
            command=self._on_auto_toggle,
        ).pack(side="left")
        tk.Label(
            top,
            text=" Run automatically in the background — no clicks required",
            bg=CARD_BG, fg=INK, font=(FONT_FAMILY, 10, "bold"),
        ).pack(side="left")

        # ── Live status badge ──
        self._sched_badge = tk.Label(
            top, text="● OFF", bg=CARD_BG,
            fg=MUTED, font=(FONT_FAMILY, 9, "bold"),
        )
        self._sched_badge.pack(side="right", padx=(0, 4))

        tk.Label(
            g,
            text="The scheduler wakes up every N minutes. After the deadline passes it sends\n"
                 "acknowledgments to submitters and reminders to everyone still missing.\n"
                 "Reminders repeat every X hours until the employee submits.",
            bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 9), justify="left",
        ).pack(anchor="w", pady=(0, 12))

        # ── Settings grid ──
        sg = tk.Frame(g, bg=CARD_BG)
        sg.pack(fill="x", pady=(0, 10))

        def _lbl(parent, text):
            tk.Label(parent, text=text, bg=CARD_BG, fg=INK,
                     font=(FONT_FAMILY, 9, "bold")).pack(anchor="w")

        def _spin(parent, var, frm, to, width=6):
            sb = tk.Spinbox(
                parent, textvariable=var, from_=frm, to=to, width=width,
                font=(FONT_FAMILY, 10), bd=1, relief="solid",
                highlightthickness=1,
                highlightbackground=BORDER, highlightcolor=ACCENT,
            )
            sb.pack(anchor="w", pady=(2, 0))
            return sb

        # Column 0 — check interval
        c0 = tk.Frame(sg, bg=CARD_BG)
        c0.pack(side="left", padx=(0, 24))
        self._auto_interval = tk.IntVar(value=int(ac.get("interval_minutes", 30)))
        _lbl(c0, "Check interval (minutes)")
        _spin(c0, self._auto_interval, 5, 240)

        # Column 1 — window start
        c1 = tk.Frame(sg, bg=CARD_BG)
        c1.pack(side="left", padx=(0, 24))
        self._auto_win_start = tk.IntVar(value=int(ac.get("window_start_hour", 7)))
        _lbl(c1, "Active from (hour, 24h)")
        _spin(c1, self._auto_win_start, 0, 23, width=5)

        # Column 2 — window end
        c2 = tk.Frame(sg, bg=CARD_BG)
        c2.pack(side="left", padx=(0, 24))
        self._auto_win_end = tk.IntVar(value=int(ac.get("window_end_hour", 20)))
        _lbl(c2, "Active until (hour, 24h)")
        _spin(c2, self._auto_win_end, 0, 23, width=5)

        # Column 3 — reminder repeat
        c3 = tk.Frame(sg, bg=CARD_BG)
        c3.pack(side="left", padx=(0, 24))
        self._auto_repeat = tk.IntVar(value=int(ac.get("reminder_repeat_hours", 24)))
        _lbl(c3, "Reminder repeat (hours)")
        _spin(c3, self._auto_repeat, 1, 168, width=5)

        # ── What to send ──
        sw = tk.Frame(g, bg=CARD_BG)
        sw.pack(fill="x", pady=(0, 10))
        self._auto_send_acks = tk.BooleanVar(value=bool(ac.get("send_acks", True)))
        self._auto_send_rems = tk.BooleanVar(value=bool(ac.get("send_reminders", True)))
        ttk.Checkbutton(sw, variable=self._auto_send_acks,
                        text=" Send acknowledgments to submitters").pack(side="left")
        ttk.Checkbutton(sw, variable=self._auto_send_rems,
                        text=" Send reminders to missing employees").pack(
            side="left", padx=(20, 0))

        # ── Buttons ──
        br = tk.Frame(g, bg=CARD_BG)
        br.pack(fill="x", pady=(4, 0))

        tk.Button(
            br, text="Save & Apply",
            font=(FONT_FAMILY, 9, "bold"),
            bg=ACCENT, fg="white", relief="flat",
            padx=14, pady=5, cursor="hand2",
            command=self._save_auto_settings,
        ).pack(side="left")

        tk.Button(
            br, text="Stop Scheduler",
            font=(FONT_FAMILY, 9),
            bg="#EDEFF6", fg=DANGER, relief="flat",
            padx=12, pady=5, cursor="hand2",
            command=self._stop_scheduler,
        ).pack(side="left", padx=(10, 0))

        # ── Last-run info ──
        self._sched_last_run = tk.Label(
            g, text="", bg=CARD_BG, fg=MUTED,
            font=(FONT_FAMILY, 8, "italic"),
        )
        self._sched_last_run.pack(anchor="w", pady=(8, 0))
        self._refresh_last_run_label()

        return card

    # ── Auto-Send helpers ──────────────────────────────────────────────
    def _on_auto_toggle(self) -> None:
        self.config.setdefault("auto_check", {})["enabled"] = \
            self._auto_enabled_var.get()
        self._save_config()
        # Restart scheduler so it picks up the new flag immediately
        email_checker.stop_auto_check()
        email_checker.schedule_auto_check(self.config_path)

    def _save_auto_settings(self) -> None:
        self.config.setdefault("auto_check", {}).update({
            "enabled":              self._auto_enabled_var.get(),
            "interval_minutes":     self._auto_interval.get(),
            "window_start_hour":    self._auto_win_start.get(),
            "window_end_hour":      self._auto_win_end.get(),
            "reminder_repeat_hours": self._auto_repeat.get(),
            "send_acks":            self._auto_send_acks.get(),
            "send_reminders":       self._auto_send_rems.get(),
        })
        self._save_config()
        # Restart with new settings
        email_checker.stop_auto_check()
        email_checker.schedule_auto_check(self.config_path)
        self._sched_last_run.configure(text="✓ Settings saved — scheduler restarted", fg=SUCCESS)
        self.root.after(4000, self._refresh_last_run_label)

    def _stop_scheduler(self) -> None:
        email_checker.stop_auto_check()
        self._sched_badge.configure(text="● OFF", fg=MUTED)
        self._sched_last_run.configure(text="Scheduler stopped manually.", fg=WARNING)

    def _refresh_last_run_label(self) -> None:
        ac = self.config.get("auto_check", {})
        last  = ac.get("last_run",    "")
        rslt  = ac.get("last_result", "")
        if last:
            self._sched_last_run.configure(
                text=f"Last run: {last}  |  Result: {rslt}", fg=MUTED,
            )
        else:
            self._sched_last_run.configure(text="Not run yet this session.", fg=MUTED)

    def _poll_scheduler_status(self) -> None:
        """Tick every 10 s — update the badge and reload last_run from disk."""
        if not self._alive:
            return
        running = email_checker.scheduler_is_running()
        enabled = self.config.get("auto_check", {}).get("enabled", False)

        if running and enabled:
            self._sched_badge.configure(text="● ACTIVE", fg=SUCCESS)
        elif running and not enabled:
            self._sched_badge.configure(text="● IDLE", fg=WARNING)
        else:
            self._sched_badge.configure(text="● OFF", fg=MUTED)

        # Reload last_run from disk (auto_runner.py may have updated it)
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                disk_cfg = json.load(f)
            ac = disk_cfg.get("auto_check", {})
            self.config["auto_check"] = ac      # keep in sync
            self._refresh_last_run_label()
        except Exception:
            pass

        self.root.after(10_000, self._poll_scheduler_status)

    # ------------------------------------------------------------------ #
    # Email Tab — Actions card
    # ------------------------------------------------------------------ #
    def _build_email_actions_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="SEND EMAILS")
        g = card.body

        tk.Label(
            g,
            text="Scan the department input folders right now and send emails accordingly.",
            bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 9),
        ).pack(anchor="w", pady=(0, 10))

        btn_row = tk.Frame(g, bg=CARD_BG)
        btn_row.pack(anchor="w")

        tk.Button(
            btn_row, text="✉ Send Reminders to Missing Depts",
            font=(FONT_FAMILY, 9, "bold"),
            bg=WARNING, fg="white", relief="flat",
            padx=14, pady=6, cursor="hand2",
            command=lambda: self._run_email_cycle(send_acks=False, send_reminders=True),
        ).pack(side="left")

        tk.Button(
            btn_row, text="✓ Send Acknowledgments to Submitters",
            font=(FONT_FAMILY, 9, "bold"),
            bg=SUCCESS, fg="white", relief="flat",
            padx=14, pady=6, cursor="hand2",
            command=lambda: self._run_email_cycle(send_acks=True, send_reminders=False),
        ).pack(side="left", padx=(10, 0))

        tk.Button(
            btn_row, text="⚡ Send All",
            font=(FONT_FAMILY, 9, "bold"),
            bg=ACCENT, fg="white", relief="flat",
            padx=14, pady=6, cursor="hand2",
            command=lambda: self._run_email_cycle(send_acks=True, send_reminders=True, force=True),
        ).pack(side="left", padx=(10, 0))

        self._email_action_status = tk.Label(
            g, text="", bg=CARD_BG, fg=MUTED,
            font=(FONT_FAMILY, 9, "italic"), wraplength=560, justify="left",
        )
        self._email_action_status.pack(anchor="w", pady=(10, 0))
        return card

    def _run_email_cycle(
        self,
        send_acks: bool = True,
        send_reminders: bool = True,
        force: bool = False,
    ) -> None:
        self._save_smtp()
        self._save_deadline()
        self._save_departments()

        if not self.config.get("email", {}).get("enabled", False):
            if not messagebox.askyesno(
                "Email Disabled",
                "Email automation is currently disabled.\n"
                "Enable it now and proceed?",
            ):
                return
            self.config["email"]["enabled"] = True
            self._email_enabled_var.set(True)
            self._save_config()

        self._email_action_status.configure(text="Scanning folders and sending emails…", fg=WARNING)
        self.root.update_idletasks()

        def _worker():
            result = email_checker.run_email_cycle(
                self.config,
                send_acks=send_acks,
                send_reminders=send_reminders,
                force=force,
            )
            def _ui():
                self._on_email_cycle_done(result)
            self.root.after(0, _ui)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_email_cycle_done(self, result: dict) -> None:
        sent   = len(result.get("emails_sent",   []))
        failed = len(result.get("emails_failed", []))
        skip   = result.get("skipped_reason")
        sub    = len(result.get("all_submitted", []))
        miss   = len(result.get("all_missing",   []))

        if skip:
            self._email_action_status.configure(text=f"ℹ {skip}", fg=MUTED)
            messagebox.showinfo("Email Cycle", skip)
            return

        summary = (
            f"Scan: {sub} submitted, {miss} missing  •  "
            f"Emails: {sent} sent, {failed} failed"
        )
        color = SUCCESS if failed == 0 else (WARNING if sent > 0 else DANGER)
        self._email_action_status.configure(text=summary, fg=color)

        errors = result.get("errors", [])
        detail = "\n".join(f"  • {e}" for e in errors) if errors else ""
        msg = f"Week: {result.get('week_label', '')}\n\n{summary}"
        if detail:
            msg += f"\n\nErrors:\n{detail}"

        if failed == 0:
            messagebox.showinfo("Email Cycle Complete", msg)
        else:
            messagebox.showwarning("Email Cycle — Some Failures", msg)

        self._refresh_submission_status()

    # ------------------------------------------------------------------ #
    # Email Tab — Submission status card
    # ------------------------------------------------------------------ #
    def _build_submission_status_card(self, parent: tk.Frame) -> tk.Frame:
        card = Card(parent, title="CURRENT SUBMISSION STATUS")
        g = card.body

        btn_row = tk.Frame(g, bg=CARD_BG)
        btn_row.pack(anchor="w", pady=(0, 10))

        tk.Button(
            btn_row, text="🔄 Refresh Status",
            font=(FONT_FAMILY, 9), bg="#EDEFF6", fg=INK,
            relief="flat", padx=12, pady=4, cursor="hand2",
            command=self._refresh_submission_status,
        ).pack(side="left")

        self._status_week_label = tk.Label(
            btn_row, text="", bg=CARD_BG, fg=MUTED, font=(FONT_FAMILY, 9, "italic"),
        )
        self._status_week_label.pack(side="left", padx=(12, 0))

        # Grid for department status rows
        self._status_grid = tk.Frame(g, bg=CARD_BG)
        self._status_grid.pack(fill="x")

        self._refresh_submission_status()
        return card

    def _refresh_submission_status(self) -> None:
        # Clear old rows
        for w in self._status_grid.winfo_children():
            w.destroy()

        scan = email_checker.check_submissions(self.config)
        self._status_week_label.configure(
            text=f"Week: {scan['week_label']}  |  "
                 f"Checked: {scan['checked_at'].replace('T', ' ')}  |  "
                 f"{scan['submit_count']}/{scan['total']} submitted",
        )

        g = self._status_grid

        if not scan["departments"]:
            tk.Label(g, text="No departments configured yet.",
                     bg=CARD_BG, fg=MUTED,
                     font=(FONT_FAMILY, 9, "italic")).pack(anchor="w", pady=4)
            return

        for dept_data in scan["departments"]:
            dept_name = dept_data["name"]
            sub_count  = len(dept_data["submitted"])
            miss_count = len(dept_data["missing"])
            total_emp  = sub_count + miss_count

            # ── Department header row ──
            dept_color = SUCCESS if miss_count == 0 and total_emp > 0 \
                         else (WARNING if sub_count > 0 else DANGER)
            dept_icon  = "✓" if miss_count == 0 and total_emp > 0 \
                         else ("◑" if sub_count > 0 else "✗")

            dept_hdr = tk.Frame(g, bg="#EDEFF6",
                                highlightbackground=BORDER, highlightthickness=1)
            dept_hdr.pack(fill="x", pady=(6, 0))

            tk.Label(
                dept_hdr,
                text=f" {dept_icon}  {dept_name}",
                bg="#EDEFF6", fg=dept_color,
                font=(FONT_FAMILY, 10, "bold"),
            ).pack(side="left", padx=(8, 0), pady=5)

            tk.Label(
                dept_hdr,
                text=f"{sub_count}/{total_emp} submitted",
                bg="#EDEFF6", fg=MUTED,
                font=(FONT_FAMILY, 9),
            ).pack(side="right", padx=(0, 12), pady=5)

            # ── Employee rows ──
            emp_frame = tk.Frame(g, bg=CARD_BG,
                                 highlightbackground=BORDER, highlightthickness=1)
            emp_frame.pack(fill="x", pady=(0, 2))

            # Column headers (once per dept)
            hdr_row = tk.Frame(emp_frame, bg="#F9FAFB")
            hdr_row.pack(fill="x")
            for txt, w in [("  Employee", 22), ("Status", 16),
                            ("File", 32), ("Submitted At", 14)]:
                tk.Label(hdr_row, text=txt, bg="#F9FAFB", fg=MUTED,
                         font=(FONT_FAMILY, 8, "bold"), width=w,
                         anchor="w").pack(side="left", padx=(0, 4), pady=3)

            # Submitted employees
            for emp in dept_data["submitted"]:
                row_f = tk.Frame(emp_frame, bg=CARD_BG)
                row_f.pack(fill="x", padx=4, pady=1)

                tk.Label(row_f, text=f"  {emp['name']}",
                         bg=CARD_BG, fg=INK,
                         font=(FONT_FAMILY, 9), width=22,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text="✓ Submitted",
                         bg=CARD_BG, fg=SUCCESS,
                         font=(FONT_FAMILY, 9, "bold"), width=16,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text=emp["matched_file"],
                         bg=CARD_BG, fg=MUTED,
                         font=(FONT_FAMILY, 8), width=32,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text=emp["modified"],
                         bg=CARD_BG, fg=MUTED,
                         font=(FONT_FAMILY, 9), anchor="w").pack(side="left")

            # Missing employees
            for emp in dept_data["missing"]:
                row_f = tk.Frame(emp_frame, bg="#FFF8F8")
                row_f.pack(fill="x", padx=4, pady=1)

                exp_note = f"  (expects: {emp['expected_file']})" \
                           if emp.get("expected_file") else ""

                tk.Label(row_f, text=f"  {emp['name']}",
                         bg="#FFF8F8", fg=INK,
                         font=(FONT_FAMILY, 9), width=22,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text="✗ Not Submitted",
                         bg="#FFF8F8", fg=DANGER,
                         font=(FONT_FAMILY, 9, "bold"), width=16,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text=f"—{exp_note}",
                         bg="#FFF8F8", fg=MUTED,
                         font=(FONT_FAMILY, 8), width=32,
                         anchor="w").pack(side="left")
                tk.Label(row_f, text="—",
                         bg="#FFF8F8", fg=MUTED,
                         font=(FONT_FAMILY, 9), anchor="w").pack(side="left")

            if total_emp == 0:
                no_row = tk.Frame(emp_frame, bg=CARD_BG)
                no_row.pack(fill="x", padx=4, pady=4)
                tk.Label(no_row,
                         text="  No employees configured for this department.",
                         bg=CARD_BG, fg=MUTED,
                         font=(FONT_FAMILY, 8, "italic")).pack(side="left")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = WeeklyReportApp(root)
        root.mainloop()
    except Exception as exc:  # noqa: BLE001
        import traceback
        try:
            messagebox.showerror("Fatal Error", traceback.format_exc())
        except Exception:
            pass  # messagebox itself failed (e.g. Tk never initialised)