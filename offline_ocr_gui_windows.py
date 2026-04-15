#!/usr/bin/env python3

from __future__ import annotations

import json
import subprocess
import threading
import sys
from pathlib import Path
from tkinter import Canvas, END, IntVar, StringVar, Text, Tk, Toplevel, filedialog, messagebox
from tkinter import ttk

import offline_batch_ocr as core
import offline_batch_ocr_windows as win_runner

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

APP_NAME = "DetailExtract OCR"
APP_VERSION = "v1.1 Fast"
APP_TAGLINE = "BY. Sarang Tohokar"
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "Sarang@9297"


class OfflineOCRGui:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("860x560")
        self.root.minsize(760, 480)

        self.selected_files: list[Path] = []
        self.runtime_root = win_runner.configure_runtime_paths()
        self.status = StringVar(value=f"Select image files to begin. Output will be saved in: {core.OUTPUT_DIR}")
        self.count_text = StringVar(value="0 files selected")
        self.progress_text = StringVar(value="Progress: 0/0")
        self.progress_value = IntVar(value=0)
        self.preview_photo = None
        self.brand_logo_photo = None
        self.monitor_duck_photo = None
        self.monitor_window: Toplevel | None = None
        self.monitor_progress_text = StringVar(value="Progress: 0/0")
        self.monitor_percent_text = StringVar(value="0%")
        self.monitor_progress_value = IntVar(value=0)
        self.monitor_log: Text | None = None
        self.theme: dict[str, str] = {}
        self.splash_window: Toplevel | None = None
        self.admin_window: Toplevel | None = None
        self.admin_logged_in = False
        self.admin_configured = False
        self.session_file = self.runtime_root / "admin_session.json"
        self._busy = False
        self._pulse_job: str | None = None
        self._pulse_on = False
        self.root.protocol("WM_DELETE_WINDOW", self.confirm_close)
        self.restore_admin_session()
        self.apply_app_icon()
        self._build_ui()
        self.show_startup_splash()

    def _current_boot_marker(self) -> str:
        if not sys.platform.startswith("win"):
            return "non-windows"
        commands = [
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-CimInstance Win32_OperatingSystem).LastBootUpTime.ToFileTimeUtc()",
            ],
            ["wmic", "os", "get", "lastbootuptime"],
        ]
        for command in commands:
            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=8,
                    check=False,
                )
            except Exception:
                continue
            output = (result.stdout or "").strip()
            if result.returncode != 0 or not output:
                continue
            cleaned = [line.strip() for line in output.splitlines() if line.strip()]
            if not cleaned:
                continue
            if "lastbootuptime" in cleaned[0].lower() and len(cleaned) > 1:
                return cleaned[1]
            return cleaned[-1]
        return "unknown-boot"

    def restore_admin_session(self) -> None:
        try:
            if not self.session_file.exists():
                return
            payload = json.loads(self.session_file.read_text(encoding="utf-8"))
            if payload.get("boot_marker") != self._current_boot_marker():
                return
            self.admin_logged_in = True
            self.admin_configured = True
            self.status.set("Admin session restored.")
        except Exception:
            return

    def save_admin_session(self) -> None:
        payload = {
            "boot_marker": self._current_boot_marker(),
        }
        try:
            self.session_file.write_text(json.dumps(payload), encoding="utf-8")
        except Exception:
            pass

    def confirm_close(self) -> None:
        if not messagebox.askyesno("Close Software", "Would you like to close?"):
            return
        try:
            self.root.destroy()
        except Exception:
            pass

    def apply_app_icon(self) -> None:
        candidates: list[Path] = []
        if getattr(sys, "frozen", False):
            candidates.append(Path(sys.executable).resolve().parent / "detailextract.ico")
            candidates.append(Path(sys.executable).resolve().parent / "assets" / "detailextract.ico")
        candidates.append(Path(__file__).resolve().parent / "assets" / "detailextract.ico")
        for icon_path in candidates:
            if icon_path.exists():
                try:
                    self.root.iconbitmap(str(icon_path))
                    return
                except Exception:
                    continue

    def _build_ui(self) -> None:
        # Dark minimal style with provided palette accents:
        # Carrot #EF8F00, Persian Blue #0038BC, Platinum #EEEEEE.
        bg_root = "#0E1423"
        bg_card = "#121B2D"
        text_title = "#EEEEEE"
        text_muted = "#B8C2DE"
        text_body = "#D5DEF6"
        mint_primary = "#EF8F00"
        mint_primary_active = "#FF9F1A"
        mint_primary_pressed = "#D67D00"
        mint_border = "#26314D"
        surface_secondary = "#0038BC"
        surface_secondary_active = "#1B4ECC"
        surface_secondary_pressed = "#002E98"
        progress_trough = "#1A2740"
        selected_row = "#213A73"
        self.theme = {
            "bg_root": bg_root,
            "bg_card": bg_card,
            "text_body": text_body,
            "mint_border": mint_border,
        }

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Root.TFrame", background=bg_root)
        style.configure("Card.TFrame", background=bg_card, relief="flat")
        style.configure("Title.TLabel", background=bg_root, foreground=text_title, font=("Segoe UI Semibold", 20, "bold"))
        style.configure("Subtitle.TLabel", background=bg_root, foreground=text_muted, font=("Segoe UI", 10))
        style.configure("SectionTitle.TLabel", background=bg_card, foreground=text_title, font=("Segoe UI Semibold", 11, "bold"))
        style.configure("Body.TLabel", background=bg_card, foreground=text_body, font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=bg_card, foreground=text_body, font=("Segoe UI", 10))
        style.configure(
            "Primary.TButton",
            font=("Segoe UI Semibold", 10, "bold"),
            padding=(12, 8),
            background=mint_primary,
            foreground="#1A1A1A",
            bordercolor=mint_border,
            lightcolor="#2C3550",
            darkcolor="#0A0F1A",
        )
        style.configure(
            "PulseA.TButton",
            font=("Segoe UI Semibold", 10, "bold"),
            padding=(12, 8),
            background=mint_primary,
            foreground="#1A1A1A",
            bordercolor=mint_border,
        )
        style.configure(
            "PulseB.TButton",
            font=("Segoe UI Semibold", 10, "bold"),
            padding=(12, 8),
            background=mint_primary_active,
            foreground="#1A1A1A",
            bordercolor=mint_border,
        )
        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            padding=(12, 8),
            background=surface_secondary,
            foreground="#EEEEEE",
            bordercolor=mint_border,
            lightcolor="#2C3550",
            darkcolor="#0A0F1A",
        )
        style.map(
            "Primary.TButton",
            background=[
                ("!disabled", mint_primary),
                ("pressed", mint_primary_pressed),
                ("active", mint_primary_active),
                ("disabled", "#D8D4CA"),
            ],
            foreground=[("!disabled", "#161616"), ("disabled", "#8B8B8B")],
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("!disabled", surface_secondary),
                ("pressed", surface_secondary_pressed),
                ("active", surface_secondary_active),
                ("disabled", "#2B3350"),
            ],
            foreground=[("!disabled", "#EEEEEE"), ("disabled", "#8F9AB9")],
        )
        style.configure(
            "Accent.Horizontal.TProgressbar",
            background=mint_primary_active,
            troughcolor=progress_trough,
            bordercolor=mint_border,
            lightcolor=mint_primary,
            darkcolor=mint_primary_active,
        )

        root_container = ttk.Frame(self.root, style="Root.TFrame", padding=20)
        root_container.pack(fill="both", expand=True)

        header = ttk.Frame(root_container, style="Root.TFrame")
        header.pack(fill="x")

        logo = Canvas(header, width=54, height=54, highlightthickness=0, bg=bg_root)
        logo.pack(side="left")
        duck_path = Path(__file__).resolve().parent / "assets" / "duck_logo.jpeg"
        if Image is not None and ImageTk is not None and duck_path.exists():
            duck = Image.open(duck_path).convert("RGB")
            duck.thumbnail((50, 50), Image.Resampling.LANCZOS)
            self.brand_logo_photo = ImageTk.PhotoImage(duck)
            logo.create_image(27, 27, image=self.brand_logo_photo)
        else:
            logo.create_oval(4, 4, 50, 50, fill="#FFFFFF", outline="#000000", width=2)
            logo.create_oval(14, 26, 22, 34, fill="#000000")
            logo.create_oval(31, 22, 39, 30, fill="#000000")
            logo.create_oval(21, 24, 35, 38, fill="#EF8F00", outline="#000000", width=2)

        title_box = ttk.Frame(header, style="Root.TFrame")
        title_box.pack(side="left", padx=(10, 0))
        ttk.Label(title_box, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_box, text=f"{APP_TAGLINE}  |  {APP_VERSION}", style="Subtitle.TLabel").pack(anchor="w", pady=(4, 0))
        ttk.Label(title_box, text="Duck Control Panel", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Label(
            title_box,
            text=f"Data folder: {self.runtime_root}",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        card = ttk.Frame(root_container, style="Card.TFrame", padding=18)
        card.pack(fill="both", expand=True, pady=(14, 0))
        card.columnconfigure(0, weight=2)
        card.columnconfigure(1, weight=1)
        card.rowconfigure(2, weight=1)

        actions = ttk.Frame(card, style="Card.TFrame")
        actions.grid(row=0, column=0, columnspan=2, sticky="ew")

        self.select_btn = ttk.Button(actions, text="📁 Select Files", style="Primary.TButton", command=self.select_files)
        self.select_btn.pack(side="left")

        self.extract_btn = ttk.Button(actions, text="🧾 Extract Details", style="Primary.TButton", command=self.extract_details)
        self.extract_btn.pack(side="left", padx=(8, 0))

        self.open_output_btn = ttk.Button(
            actions,
            text="📂 Open Output Folder",
            style="Secondary.TButton",
            command=self.open_output,
        )
        self.open_output_btn.pack(side="left", padx=(8, 0))

        self.clear_btn = ttk.Button(actions, text="🧹 Clear", style="Secondary.TButton", command=self.clear_files)
        self.clear_btn.pack(side="right")

        info_row = ttk.Frame(card, style="Card.TFrame")
        info_row.grid(row=1, column=0, sticky="ew", pady=(14, 8))
        ttk.Label(info_row, text="Selected files", style="SectionTitle.TLabel").pack(side="left")
        ttk.Label(info_row, textvariable=self.count_text, style="Body.TLabel").pack(side="right")

        list_wrap = ttk.Frame(card, style="Card.TFrame")
        list_wrap.grid(row=2, column=0, sticky="nsew")
        list_wrap.columnconfigure(0, weight=1)
        list_wrap.rowconfigure(0, weight=1)

        self.file_list = ttk.Treeview(
            list_wrap,
            columns=("path",),
            show="headings",
            selectmode="browse",
            height=12,
        )
        self.file_list.configure(style="Files.Treeview")
        style.configure(
            "Files.Treeview",
            background="#0F1628",
            fieldbackground="#0F1628",
            foreground=text_body,
            rowheight=28,
            bordercolor=mint_border,
            borderwidth=1,
        )
        style.configure(
            "Files.Treeview.Heading",
            background=surface_secondary,
            foreground=text_title,
            font=("Segoe UI Semibold", 10, "bold"),
        )
        style.map("Files.Treeview", background=[("selected", selected_row)], foreground=[("selected", "#EEEEEE")])
        self.file_list.heading("path", text="File path")
        self.file_list.column("path", anchor="w", stretch=True, width=740)
        self.file_list.grid(row=0, column=0, sticky="nsew")
        self.file_list.bind("<<TreeviewSelect>>", self.on_file_select)

        scrollbar = ttk.Scrollbar(list_wrap, orient="vertical", command=self.file_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list.configure(yscrollcommand=scrollbar.set)

        preview_wrap = ttk.Frame(card, style="Card.TFrame")
        preview_wrap.grid(row=1, column=1, rowspan=2, sticky="nsew", padx=(14, 0))
        preview_wrap.columnconfigure(0, weight=1)
        preview_wrap.rowconfigure(1, weight=1)
        ttk.Label(preview_wrap, text="Image preview", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.preview_label = ttk.Label(
            preview_wrap,
            text="No preview",
            style="Body.TLabel",
            anchor="center",
            justify="center",
        )
        self.preview_label.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        status_block = ttk.Frame(card, style="Card.TFrame")
        status_block.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.progress = ttk.Progressbar(
            status_block,
            mode="determinate",
            style="Accent.Horizontal.TProgressbar",
            variable=self.progress_value,
            maximum=100,
        )
        self.progress.pack(fill="x")
        ttk.Label(status_block, textvariable=self.progress_text, style="Body.TLabel").pack(anchor="w", pady=(6, 0))
        ttk.Label(status_block, textvariable=self.status, style="Status.TLabel", wraplength=790).pack(anchor="w", pady=(8, 0))
        self.setup_button_interactions()
        self.extract_btn.config(state="disabled")
        self.status.set("Login required before using software.")

    def setup_button_interactions(self) -> None:
        for button in (self.select_btn, self.extract_btn, self.open_output_btn, self.clear_btn):
            try:
                button.configure(cursor="hand2")
            except Exception:
                pass
            button.bind("<Enter>", lambda _e, btn=button: btn.state(["active"]))
            button.bind("<Leave>", lambda _e, btn=button: btn.state(["!active"]))

    def start_extract_pulse(self) -> None:
        if self._busy:
            return
        if self._pulse_job is not None:
            self.root.after_cancel(self._pulse_job)
            self._pulse_job = None
        self._pulse_on = False

        def tick() -> None:
            if self._busy:
                self.extract_btn.configure(style="Primary.TButton")
                self._pulse_job = None
                return
            self._pulse_on = not self._pulse_on
            self.extract_btn.configure(style="PulseA.TButton" if self._pulse_on else "PulseB.TButton")
            self._pulse_job = self.root.after(420, tick)

        tick()

    def stop_extract_pulse(self) -> None:
        if self._pulse_job is not None:
            self.root.after_cancel(self._pulse_job)
            self._pulse_job = None
        self.extract_btn.configure(style="Primary.TButton")

    def select_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select OCR image files",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        self.selected_files = [Path(p) for p in paths]
        self.refresh_file_list()
        self.show_preview(self.selected_files[0])
        self.progress_value.set(0)
        self.progress_text.set(f"Progress: 0/{len(self.selected_files)}")
        self.status.set("Files selected. Click Extract Details to start.")
        self.update_extract_button_state()
        self.start_extract_pulse()

    def refresh_file_list(self) -> None:
        for item in self.file_list.get_children():
            self.file_list.delete(item)
        for file_path in self.selected_files:
            self.file_list.insert("", END, values=(str(file_path),))
        count = len(self.selected_files)
        noun = "file" if count == 1 else "files"
        self.count_text.set(f"{count} {noun} selected")

    def clear_files(self) -> None:
        if self._busy:
            messagebox.showinfo("Please wait", "Extraction is running. Wait for it to finish before clearing files.")
            return
        self.selected_files = []
        for item in self.file_list.selection():
            self.file_list.selection_remove(item)
        for item in self.file_list.get_children():
            self.file_list.delete(item)
        self.count_text.set("0 files selected")
        self.preview_label.configure(image="", text="No preview")
        self.preview_photo = None
        self.progress_value.set(0)
        self.progress_text.set("Progress: 0/0")
        self.monitor_progress_value.set(0)
        self.monitor_progress_text.set("Progress: 0/0")
        self.monitor_percent_text.set("0%")
        if self.monitor_log is not None:
            self.monitor_log.delete("1.0", END)
        self.status.set("Selection cleared.")
        self.stop_extract_pulse()
        self.update_extract_button_state()
        self.root.update_idletasks()

    def set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.select_btn.config(state=state)
        self.clear_btn.config(state=state)
        self.update_extract_button_state(force_disable=busy)
        if busy:
            self.stop_extract_pulse()
        elif self.selected_files:
            self.start_extract_pulse()

    def update_extract_button_state(self, force_disable: bool = False) -> None:
        allowed = (
            (not force_disable)
            and self.admin_configured
            and bool(self.selected_files)
        )
        self.extract_btn.config(state="normal" if allowed else "disabled")

    def show_startup_splash(self) -> None:
        self.splash_window = Toplevel(self.root)
        self.splash_window.overrideredirect(True)
        self.splash_window.configure(bg="#0E1423")
        self.splash_window.geometry("420x240")
        self.splash_window.attributes("-topmost", True)

        frame = ttk.Frame(self.splash_window, style="Card.TFrame", padding=18)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        duck_canvas = Canvas(frame, width=120, height=80, highlightthickness=0, bg="#121B2D")
        duck_canvas.grid(row=0, column=0, sticky="w", padx=(0, 12))
        duck_path = Path(__file__).resolve().parent / "assets" / "duck_logo.jpeg"
        duck_item = None
        if Image is not None and ImageTk is not None and duck_path.exists():
            duck = Image.open(duck_path).convert("RGB")
            duck.thumbnail((66, 66), Image.Resampling.LANCZOS)
            self.monitor_duck_photo = ImageTk.PhotoImage(duck)
            duck_item = duck_canvas.create_image(30, 40, image=self.monitor_duck_photo)
        else:
            duck_item = duck_canvas.create_text(30, 40, text="🦆", font=("Segoe UI Emoji", 30), fill="#EEEEEE")

        title = ttk.Label(frame, text=APP_NAME, style="Title.TLabel")
        title.grid(row=0, column=1, sticky="sw")
        ttk.Label(frame, text=f"{APP_TAGLINE}  |  {APP_VERSION}", style="Subtitle.TLabel").grid(row=1, column=1, sticky="nw", pady=(2, 0))
        anim = ttk.Label(frame, text="Loading control panel...", style="Body.TLabel")
        anim.grid(row=2, column=1, sticky="nw", pady=(12, 0))

        self.root.withdraw()

        def finish() -> None:
            if self.splash_window is not None and self.splash_window.winfo_exists():
                self.splash_window.destroy()
            if self.admin_configured:
                self.root.deiconify()
                self.update_extract_button_state()
                return
            self.open_admin_portal(startup_gate=True)

        if duck_item is not None:
            steps = {"i": 0, "dir": 1}

            def animate() -> None:
                if self.splash_window is None or not self.splash_window.winfo_exists():
                    return
                x, y = duck_canvas.coords(duck_item)
                if x >= 88:
                    steps["dir"] = -1
                elif x <= 30:
                    steps["dir"] = 1
                duck_canvas.coords(duck_item, x + (5 * steps["dir"]), y)
                anim.configure(text=f"Loading control panel{'.' * ((steps['i'] % 3) + 1)}")
                steps["i"] += 1
                self.root.after(110, animate)

            animate()

        self.root.after(2000, finish)

    def open_admin_portal(self, startup_gate: bool = False) -> None:
        if self.admin_window is not None and self.admin_window.winfo_exists():
            self.admin_window.lift()
            return

        self.admin_window = Toplevel(self.root)
        self.admin_window.title("Admin Portal")
        self.admin_window.geometry("420x280")
        self.admin_window.resizable(False, False)
        self.admin_window.configure(bg=self.theme.get("bg_root", "#0E1423"))

        panel = ttk.Frame(self.admin_window, style="Card.TFrame", padding=16)
        panel.pack(fill="both", expand=True)
        ttk.Label(panel, text="🔐 Admin Portal", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(panel, text=f"Publisher: BY. Sarang Tohokar  |  {APP_VERSION}", style="Body.TLabel").pack(anchor="w", pady=(4, 10))

        form = ttk.Frame(panel, style="Card.TFrame")
        form.pack(fill="x")
        ttk.Label(form, text="Username", style="Body.TLabel").grid(row=0, column=0, sticky="w")
        user_var = StringVar(value=ADMIN_USERNAME)
        user_entry = ttk.Entry(form, textvariable=user_var, width=28)
        user_entry.grid(row=1, column=0, sticky="ew", pady=(2, 8))

        ttk.Label(form, text="Password", style="Body.TLabel").grid(row=2, column=0, sticky="w")
        pass_var = StringVar(value="")
        pass_entry = ttk.Entry(form, textvariable=pass_var, width=28, show="*")
        pass_entry.grid(row=3, column=0, sticky="ew", pady=(2, 8))

        msg_var = StringVar(value="")
        ttk.Label(form, textvariable=msg_var, style="Body.TLabel").grid(row=4, column=0, sticky="w")

        def apply_admin() -> None:
            username = user_var.get().strip()
            password = pass_var.get()
            if username != ADMIN_USERNAME or password != ADMIN_PASSWORD:
                msg_var.set("Invalid username or password.")
                return

            self.admin_logged_in = True
            self.admin_configured = True
            self.status.set("Admin login successful.")
            self.save_admin_session()
            self.update_extract_button_state()
            self.admin_window.destroy()
            if startup_gate:
                self.root.deiconify()

        if startup_gate:
            def close_block() -> None:
                self.confirm_close()

            self.admin_window.protocol("WM_DELETE_WINDOW", close_block)

        ttk.Button(panel, text="✅ Apply Admin Settings", style="Primary.TButton", command=apply_admin).pack(anchor="e", pady=(8, 0))

    def _eligible_files(self) -> list[Path]:
        return [path for path in self.selected_files if path.suffix.lower() in core.SUPPORTED_EXTENSIONS]

    def show_preview(self, image_path: Path) -> None:
        if Image is None or ImageTk is None:
            self.preview_label.configure(text=f"Preview unavailable\n{image_path.name}", image="")
            self.preview_photo = None
            return
        try:
            image = Image.open(image_path)
            image.thumbnail((260, 260))
            self.preview_photo = ImageTk.PhotoImage(image)
            self.preview_label.configure(image=self.preview_photo, text="")
        except Exception:
            self.preview_label.configure(text=f"Preview unavailable\n{image_path.name}", image="")
            self.preview_photo = None

    def on_file_select(self, _event: object) -> None:
        selected = self.file_list.selection()
        if not selected:
            return
        values = self.file_list.item(selected[0], "values")
        if not values:
            return
        path = Path(values[0])
        if path.exists():
            self.show_preview(path)

    def ensure_monitor_window(self) -> None:
        if self.monitor_window is not None and self.monitor_window.winfo_exists():
            return

        self.monitor_window = Toplevel(self.root)
        self.monitor_window.title("Extraction Monitor")
        self.monitor_window.geometry("720x460")
        self.monitor_window.minsize(620, 380)
        self.monitor_window.configure(bg=self.theme.get("bg_root", "#F3F1EC"))

        container = ttk.Frame(self.monitor_window, style="Card.TFrame", padding=14)
        container.pack(fill="both", expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(2, weight=1)

        header_row = ttk.Frame(container, style="Card.TFrame")
        header_row.grid(row=0, column=0, sticky="ew")
        header_row.columnconfigure(1, weight=1)
        duck_label = ttk.Label(header_row, style="Body.TLabel", text="")
        duck_label.grid(row=0, column=0, sticky="w", padx=(0, 8))
        duck_path = Path(__file__).resolve().parent / "assets" / "duck_logo.jpeg"
        if Image is not None and ImageTk is not None and duck_path.exists():
            duck = Image.open(duck_path).convert("RGB")
            duck.thumbnail((42, 42), Image.Resampling.LANCZOS)
            self.monitor_duck_photo = ImageTk.PhotoImage(duck)
            duck_label.configure(image=self.monitor_duck_photo)
        else:
            duck_label.configure(text="DUCK")
        title_wrap = ttk.Frame(header_row, style="Card.TFrame")
        title_wrap.grid(row=0, column=1, sticky="w")
        ttk.Label(title_wrap, text="Live Extraction Status", style="SectionTitle.TLabel").pack(anchor="w")
        ttk.Label(title_wrap, text="Control Panel Publisher: Sarang Tohokar", style="Body.TLabel").pack(anchor="w")

        row = ttk.Frame(container, style="Card.TFrame")
        row.grid(row=1, column=0, sticky="ew", pady=(10, 8))
        row.columnconfigure(0, weight=1)
        ttk.Label(row, textvariable=self.monitor_progress_text, style="Body.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(row, textvariable=self.monitor_percent_text, style="Body.TLabel").grid(row=0, column=1, sticky="e")
        ttk.Progressbar(
            row,
            mode="determinate",
            style="Accent.Horizontal.TProgressbar",
            variable=self.monitor_progress_value,
            maximum=100,
        ).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        log_frame = ttk.Frame(container, style="Card.TFrame")
        log_frame.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.monitor_log = Text(
            log_frame,
            bg="#F8F7F2",
            fg=self.theme.get("text_body", "#3E3E39"),
            insertbackground=self.theme.get("text_body", "#3E3E39"),
            relief="flat",
            wrap="word",
            font=("Consolas", 10),
            highlightthickness=1,
            highlightbackground=self.theme.get("mint_border", "#CFCBC2"),
            highlightcolor=self.theme.get("mint_border", "#CFCBC2"),
        )
        self.monitor_log.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.monitor_log.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.monitor_log.configure(yscrollcommand=yscroll.set)

    def append_monitor_log(self, message: str) -> None:
        self.ensure_monitor_window()
        if self.monitor_log is None:
            return
        self.monitor_log.insert(END, message + "\n")
        self.monitor_log.see(END)

    def extract_details(self) -> None:
        if not self.admin_configured:
            messagebox.showwarning("Admin required", "Login first to use the software.")
            return
        if not self.selected_files:
            messagebox.showwarning("No files selected", "Select at least one image file first.")
            return

        missing = [str(path) for path in self.selected_files if not path.exists()]
        if missing:
            messagebox.showerror("Missing files", "Some selected files no longer exist. Re-select files.")
            return

        self.set_busy(True)
        self.status.set("Starting extraction...")
        self.ensure_monitor_window()
        self.monitor_progress_value.set(0)
        self.monitor_progress_text.set("Progress: 0/0")
        self.monitor_percent_text.set("0%")
        if self.monitor_log is not None:
            self.monitor_log.delete("1.0", END)
            self.monitor_log.insert(END, "Starting extraction...\n")

        def worker() -> None:
            try:
                core.ensure_directories()
                files = self._eligible_files()
                if not files:
                    self.root.after(
                        0,
                        lambda: self.fail_extraction("No supported image files found in your selection."),
                    )
                    return
                tesseract_path = win_runner.find_windows_tesseract()
                if not tesseract_path:
                    self.root.after(
                        0,
                        lambda: self.fail_extraction(
                            "Tesseract is missing from the installed app package. Rebuild installer and try again."
                        ),
                    )
                    return
                self.root.after(0, lambda: self.status.set(f"Running OCR on {len(files)} file(s)..."))
                self.root.after(0, lambda: self.progress_text.set(f"Progress: 0/{len(files)}"))
                self.root.after(0, lambda: self.progress_value.set(0))
                self.root.after(0, lambda: self.monitor_progress_text.set(f"Progress: 0/{len(files)}"))
                self.root.after(0, lambda: self.monitor_percent_text.set("0%"))
                self.root.after(0, lambda: self.append_monitor_log(f"Running OCR on {len(files)} file(s)"))

                def on_progress(done: int, total: int, name: str) -> None:
                    percent = int((done / total) * 100) if total else 0
                    self.root.after(0, lambda: self.progress_value.set(percent))
                    self.root.after(0, lambda: self.progress_text.set(f"Progress: {done}/{total} ({percent}%)"))
                    self.root.after(0, lambda: self.status.set(f"Processed {done}/{total}: {name}"))
                    self.root.after(0, lambda: self.monitor_progress_value.set(percent))
                    self.root.after(0, lambda: self.monitor_progress_text.set(f"Progress: {done}/{total}"))
                    self.root.after(0, lambda: self.monitor_percent_text.set(f"{percent}%"))

                def on_message(message: str) -> None:
                    if message.startswith("Starting "):
                        import re

                        match = re.search(r"Starting (\d+)/(\d+):", message)
                        if match:
                            current = int(match.group(1))
                            total = int(match.group(2))
                            percent = max(1, int(((current - 1) / total) * 100)) if total else 0
                            self.root.after(0, lambda: self.progress_value.set(percent))
                            self.root.after(0, lambda: self.progress_text.set(f"Working: {current}/{total}"))
                            self.root.after(0, lambda: self.monitor_progress_value.set(percent))
                            self.root.after(0, lambda: self.monitor_progress_text.set(f"Working: {current}/{total}"))
                            self.root.after(0, lambda: self.monitor_percent_text.set(f"{percent}%"))
                    self.root.after(0, lambda: self.append_monitor_log(message))

                rc = win_runner.process_images(
                    files,
                    tesseract_path,
                    progress_callback=on_progress,
                    message_callback=on_message,
                )
                self.root.after(0, lambda: self.finish_extraction(rc))
            except Exception as error:
                self.root.after(0, lambda: self.fail_extraction(str(error)))

        threading.Thread(target=worker, daemon=True).start()

    def finish_extraction(self, rc: int) -> None:
        self.set_busy(False)
        if rc == 0:
            self.progress_value.set(100)
            self.progress_text.set("Progress: complete")
            self.monitor_progress_value.set(100)
            self.monitor_percent_text.set("100%")
            self.monitor_progress_text.set("Progress: complete")
            self.append_monitor_log("Extraction finished successfully.")
            self.status.set(f"Extraction complete. Output saved in: {core.OUTPUT_DIR}")
            messagebox.showinfo("Completed", f"Extraction completed.\n\nOutput folder:\n{core.OUTPUT_DIR}")
        else:
            self.progress_text.set("Progress: failed")
            self.monitor_progress_text.set("Progress: failed")
            self.append_monitor_log("Extraction failed.")
            self.status.set("Extraction failed. Check Tesseract setup and try again.")
            detail = win_runner.get_last_failure_reason()
            error_message = (
                "OCR failed. Make sure Tesseract is bundled in the app package and try again."
                + (f"\n\nDetails:\n{detail}" if detail else "")
            )
            messagebox.showerror(
                "Extraction failed",
                error_message,
            )

    def fail_extraction(self, message: str) -> None:
        self.set_busy(False)
        self.progress_text.set("Progress: error")
        self.monitor_progress_text.set("Progress: error")
        self.append_monitor_log(f"Error: {message}")
        self.status.set("Unexpected error during extraction.")
        messagebox.showerror("Error", message)

    def open_output(self) -> None:
        core.ensure_directories()
        try:
            import os

            os.startfile(str(core.OUTPUT_DIR))
        except Exception as error:
            messagebox.showerror("Cannot open folder", str(error))


def main() -> int:
    root = Tk()
    OfflineOCRGui(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
