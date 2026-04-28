# app.py — TE Template Filler (CustomTkinter UI)
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog

from engine import fill_plant, default_out_path
from plants import REGISTRY, find_templates, match_filename

# ── Appearance ────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ── Config persistence ────────────────────────────────────────────────────────
_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

def _load_config() -> dict:
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def _save_config(cfg: dict) -> None:
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


# ── Main App ──────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    SIDEBAR_W = 200
    ACCENT    = "#1f6feb"

    def __init__(self):
        super().__init__()
        self.title("TE Template Filler — Mexico")
        self.geometry("1000x640")
        self.minsize(900, 580)
        self.resizable(True, True)

        self._cfg = _load_config()
        self._running = False
        self._selected_plant: str | None = None

        self._build_layout()
        self._restore_state()

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=self.SIDEBAR_W, corner_radius=0)
        sb.grid(row=0, column=0, sticky="nsew")
        sb.grid_propagate(False)
        sb.grid_rowconfigure(99, weight=1)

        ctk.CTkLabel(
            sb, text="TE Filler", font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=0, column=0, padx=16, pady=(20, 4), sticky="w")
        ctk.CTkLabel(
            sb, text="Mexico Sites", font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=1, column=0, padx=16, pady=(0, 16), sticky="w")

        ctk.CTkLabel(
            sb, text="SELECT PLANT", font=ctk.CTkFont(size=10, weight="bold"),
            text_color="gray"
        ).grid(row=2, column=0, padx=16, pady=(0, 4), sticky="w")

        plants = sorted(REGISTRY.keys())
        self._plant_btns: dict[str, ctk.CTkButton] = {}
        for i, key in enumerate(plants):
            btn = ctk.CTkButton(
                sb,
                text=key.replace("_", " ").title(),
                anchor="w",
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray80", "gray25"),
                corner_radius=6,
                command=lambda k=key: self._select_plant(k),
            )
            btn.grid(row=3 + i, column=0, padx=8, pady=2, sticky="ew")
            self._plant_btns[key] = btn

        ctk.CTkLabel(sb, text="", text_color="gray").grid(
            row=98, column=0, padx=16, pady=4, sticky="w"
        )
        self._theme_switch = ctk.CTkSwitch(
            sb, text="Light mode",
            command=self._toggle_theme,
            onvalue="light", offvalue="dark"
        )
        self._theme_switch.grid(row=100, column=0, padx=16, pady=(4, 20), sticky="w")

    def _build_main(self):
        main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=16, pady=16)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)

        self._tabs = ctk.CTkTabview(main)
        self._tabs.grid(row=0, column=0, sticky="nsew")
        self._tabs.add("Single Plant")
        self._tabs.add("Run All")

        self._build_single_tab(self._tabs.tab("Single Plant"))
        self._build_batch_tab(self._tabs.tab("Run All"))

    # ── Single-plant tab ──────────────────────────────────────────────────────
    def _build_single_tab(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(7, weight=1)

        self._title_label = ctk.CTkLabel(
            parent,
            text="Select a plant from the sidebar",
            font=ctk.CTkFont(size=18, weight="bold"),
        )
        self._title_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(4, 16))

        def file_row(row, label, var, browse_cmd):
            ctk.CTkLabel(parent, text=label, anchor="w", width=140).grid(
                row=row, column=0, sticky="w", pady=6
            )
            entry = ctk.CTkEntry(parent, textvariable=var, width=500)
            entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=6)
            ctk.CTkButton(parent, text="Browse", width=80, command=browse_cmd).grid(
                row=row, column=2, pady=6
            )

        self._db_var    = ctk.StringVar()
        self._blank_var = ctk.StringVar()
        self._out_var   = ctk.StringVar()

        file_row(1, "Database (.xlsm):",      self._db_var,    self._browse_db)
        file_row(2, "Blank Template (.xlsm):", self._blank_var, self._browse_blank)
        file_row(3, "Output (Filled TE):",     self._out_var,   self._browse_out)

        opts = ctk.CTkFrame(parent, fg_color="transparent")
        opts.grid(row=4, column=0, columnspan=3, sticky="w", pady=(4, 12))
        self._visible_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts, text="Show Excel while running (slower)",
            variable=self._visible_var
        ).pack(side="left")

        self._run_btn = ctk.CTkButton(
            parent,
            text="Run Fill",
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._run,
            state="disabled",
        )
        self._run_btn.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))

        self._progress = ctk.CTkProgressBar(parent, mode="indeterminate")
        self._progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self._progress.set(0)

        log_frame = ctk.CTkFrame(parent)
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew")
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self._log = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(family="Courier", size=12))
        self._log.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self._log.configure(state="disabled")

        ctk.CTkLabel(
            parent, text="Log", font=ctk.CTkFont(size=12, weight="bold"), anchor="w"
        ).grid(row=7, column=0, columnspan=3, sticky="nw", padx=4, pady=(4, 0))

    # ── Batch tab (Run All) ───────────────────────────────────────────────────
    def _build_batch_tab(self, parent):
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(7, weight=1)

        ctk.CTkLabel(
            parent,
            text="Run All Plants at Once",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(4, 16))

        def file_row(row, label, var, browse_cmd):
            ctk.CTkLabel(parent, text=label, anchor="w", width=140).grid(
                row=row, column=0, sticky="w", pady=6
            )
            entry = ctk.CTkEntry(parent, textvariable=var, width=500)
            entry.grid(row=row, column=1, sticky="ew", padx=(8, 8), pady=6)
            ctk.CTkButton(parent, text="Browse", width=80, command=browse_cmd).grid(
                row=row, column=2, pady=6
            )

        self._batch_db_var     = ctk.StringVar()
        self._batch_folder_var = ctk.StringVar()

        file_row(1, "Database (.xlsm):",   self._batch_db_var,     self._browse_batch_db)
        file_row(2, "Templates Folder:",   self._batch_folder_var, self._browse_batch_folder)

        opts = ctk.CTkFrame(parent, fg_color="transparent")
        opts.grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 4))

        self._batch_visible_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts, text="Show Excel while running (slower)",
            variable=self._batch_visible_var
        ).pack(side="left", padx=(0, 20))

        ctk.CTkButton(
            opts, text="Preview Matches", width=140,
            command=self._preview_matches
        ).pack(side="left")

        self._run_all_btn = ctk.CTkButton(
            parent,
            text="Run All",
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._run_all,
        )
        self._run_all_btn.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 8))

        self._batch_progress = ctk.CTkProgressBar(parent)
        self._batch_progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self._batch_progress.set(0)

        # Matched-files label
        self._batch_status_label = ctk.CTkLabel(
            parent, text="", font=ctk.CTkFont(size=11), text_color="gray", anchor="w"
        )
        self._batch_status_label.grid(row=6, column=0, columnspan=3, sticky="w", padx=4)

        log_frame = ctk.CTkFrame(parent)
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew")
        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        self._batch_log = ctk.CTkTextbox(
            log_frame, font=ctk.CTkFont(family="Courier", size=12)
        )
        self._batch_log.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        self._batch_log.configure(state="disabled")

        ctk.CTkLabel(
            parent, text="Log", font=ctk.CTkFont(size=12, weight="bold"), anchor="w"
        ).grid(row=7, column=0, columnspan=3, sticky="nw", padx=4, pady=(4, 0))

    # ── State / restore ───────────────────────────────────────────────────────
    def _restore_state(self):
        db = self._cfg.get("last_db", "")
        if db and os.path.exists(db):
            self._db_var.set(db)
            self._batch_db_var.set(db)

        folder = self._cfg.get("last_folder", "")
        if folder and os.path.isdir(folder):
            self._batch_folder_var.set(folder)

        last_plant = self._cfg.get("last_plant")
        if last_plant and last_plant in REGISTRY:
            self._select_plant(last_plant)

    # ── Plant selection ───────────────────────────────────────────────────────
    def _select_plant(self, key: str):
        for k, b in self._plant_btns.items():
            b.configure(fg_color="transparent")

        self._plant_btns[key].configure(fg_color=self.ACCENT)
        self._selected_plant = key
        self._title_label.configure(text=f"Plant: {key.replace('_', ' ').title()}")
        self._run_btn.configure(state="normal")
        self._tabs.set("Single Plant")

        self._cfg["last_plant"] = key
        _save_config(self._cfg)

    # ── File pickers — single plant ───────────────────────────────────────────
    def _browse_db(self):
        path = filedialog.askopenfilename(
            title="Select Database (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._db_var.set(path)
            self._batch_db_var.set(path)   # keep both in sync
            self._cfg["last_db"] = path
            _save_config(self._cfg)

    def _browse_blank(self):
        path = filedialog.askopenfilename(
            title="Select Blank TE Template (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._blank_var.set(path)
            self._out_var.set(default_out_path(path))

    def _browse_out(self):
        path = filedialog.asksaveasfilename(
            title="Save Filled TE as…",
            defaultextension=".xlsm",
            filetypes=[("Excel Macro-Enabled", "*.xlsm")],
        )
        if path:
            self._out_var.set(path)

    # ── File pickers — batch ──────────────────────────────────────────────────
    def _browse_batch_db(self):
        path = filedialog.askopenfilename(
            title="Select Database (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._batch_db_var.set(path)
            self._db_var.set(path)         # keep both in sync
            self._cfg["last_db"] = path
            _save_config(self._cfg)

    def _browse_batch_folder(self):
        folder = filedialog.askdirectory(title="Select Templates Folder")
        if folder:
            self._batch_folder_var.set(folder)
            self._cfg["last_folder"] = folder
            _save_config(self._cfg)
            self._preview_matches()        # auto-preview on folder pick

    # ── Preview matches ───────────────────────────────────────────────────────
    def _preview_matches(self):
        folder = self._batch_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self._batch_status_label.configure(
                text="No folder selected.", text_color="gray"
            )
            return

        matches = find_templates(folder)
        if not matches:
            self._batch_status_label.configure(
                text="No matching .xlsm templates found in folder.", text_color="#f85149"
            )
            self._batch_log_clear()
            return

        self._batch_status_label.configure(
            text=f"{len(matches)} template(s) matched — ready to run.",
            text_color="#3fb950"
        )
        self._batch_log_clear()
        self._batch_log_write(f"Matched {len(matches)} template(s):", "info")
        for key, fpath in matches:
            self._batch_log_write(
                f"  {key:<22}  {os.path.basename(fpath)}", ""
            )

    # ── Logging — single plant ────────────────────────────────────────────────
    def _log_write(self, msg: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.configure(state="normal")
        self._log.insert("end", f"[{ts}] {msg}\n")
        if tag:
            start = self._log.index("end - 1 lines linestart")
            self._log.tag_add(tag, start, "end - 1c")
            self._log.tag_config("error",   foreground="#f85149")
            self._log.tag_config("success", foreground="#3fb950")
            self._log.tag_config("info",    foreground="#58a6ff")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ── Logging — batch ───────────────────────────────────────────────────────
    def _batch_log_write(self, msg: str, tag: str = ""):
        ts = datetime.now().strftime("%H:%M:%S")
        self._batch_log.configure(state="normal")
        self._batch_log.insert("end", f"[{ts}] {msg}\n")
        if tag:
            start = self._batch_log.index("end - 1 lines linestart")
            self._batch_log.tag_add(tag, start, "end - 1c")
            self._batch_log.tag_config("error",   foreground="#f85149")
            self._batch_log.tag_config("success", foreground="#3fb950")
            self._batch_log.tag_config("info",    foreground="#58a6ff")
        self._batch_log.see("end")
        self._batch_log.configure(state="disabled")

    def _batch_log_clear(self):
        self._batch_log.configure(state="normal")
        self._batch_log.delete("1.0", "end")
        self._batch_log.configure(state="disabled")

    # ── Run — single plant ────────────────────────────────────────────────────
    def _run(self):
        if self._running:
            return

        db        = self._db_var.get().strip()
        blank     = self._blank_var.get().strip()
        out       = self._out_var.get().strip()
        plant_key = self._selected_plant

        if not db or not os.path.exists(db):
            self._log_write("ERROR: Please select a valid Database (.xlsm) file.", "error")
            return
        if not blank or not os.path.exists(blank):
            self._log_write("ERROR: Please select a valid Blank Template (.xlsm) file.", "error")
            return
        if not plant_key or plant_key not in REGISTRY:
            self._log_write("ERROR: Please select a plant from the sidebar.", "error")
            return
        if not out:
            out = default_out_path(blank)
            self._out_var.set(out)

        self._log_clear()
        self._log_write(f"Starting fill for: {plant_key}", "info")
        self._log_write(f"DB:       {db}", "info")
        self._log_write(f"Template: {blank}", "info")
        self._log_write(f"Output:   {out}", "info")
        self._log_write("Working…", "info")

        self._running = True
        self._run_btn.configure(state="disabled", text="Running…")
        self._progress.configure(mode="indeterminate")
        self._progress.start()

        def _worker():
            try:
                fill_plant(db, blank, out, REGISTRY[plant_key],
                           visible=self._visible_var.get())
                self.after(0, self._on_success, out)
            except Exception as exc:
                self.after(0, self._on_error, str(exc), traceback.format_exc())

        threading.Thread(target=_worker, daemon=True).start()

    def _on_success(self, out_path: str):
        self._progress.stop()
        self._progress.set(1)
        self._run_btn.configure(state="normal", text="Run Fill")
        self._running = False
        self._log_write(f"Done!  Saved to: {out_path}", "success")

    def _on_error(self, short: str, full: str):
        self._progress.stop()
        self._progress.set(0)
        self._run_btn.configure(state="normal", text="Run Fill")
        self._running = False
        self._log_write(f"ERROR: {short}", "error")
        self._log.configure(state="normal")
        self._log.insert("end", full + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    # ── Run All ───────────────────────────────────────────────────────────────
    def _run_all(self):
        if self._running:
            return

        db     = self._batch_db_var.get().strip()
        folder = self._batch_folder_var.get().strip()

        if not db or not os.path.exists(db):
            self._batch_log_write("ERROR: Please select a valid Database (.xlsm) file.", "error")
            return
        if not folder or not os.path.isdir(folder):
            self._batch_log_write("ERROR: Please select a valid templates folder.", "error")
            return

        matches = find_templates(folder)
        if not matches:
            self._batch_log_write(
                "No matching .xlsm templates found in the selected folder.", "error"
            )
            return

        self._batch_log_clear()
        self._batch_log_write(f"Starting batch run — {len(matches)} plant(s):", "info")
        for key, fpath in matches:
            self._batch_log_write(f"  {key}: {os.path.basename(fpath)}", "info")
        self._batch_log_write("", "")

        self._running = True
        self._run_all_btn.configure(state="disabled", text="Running…")
        self._batch_progress.set(0)

        total = len(matches)

        def _worker():
            done = 0
            errors: list[tuple[str, str]] = []

            for i, (key, fpath) in enumerate(matches, start=1):
                self.after(0, self._batch_log_write,
                           f"[{i}/{total}] {key} …", "info")
                try:
                    fill_plant(db, fpath, fpath, REGISTRY[key],
                               visible=self._batch_visible_var.get())
                    done += 1
                    self.after(0, self._batch_log_write,
                               f"        → saved: {os.path.basename(fpath)}", "success")
                except Exception as exc:
                    errors.append((key, str(exc)))
                    self.after(0, self._batch_log_write,
                               f"        ERROR: {exc}", "error")
                    self.after(0, self._batch_log_write,
                               traceback.format_exc(), "error")

                self.after(0, self._batch_progress.set, i / total)

            self.after(0, self._on_batch_done, done, total, errors)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_batch_done(self, done: int, total: int, errors: list):
        self._running = False
        self._run_all_btn.configure(state="normal", text="Run All")
        self._batch_log_write("", "")
        if errors:
            self._batch_log_write(
                f"Completed {done}/{total} plants.  {len(errors)} error(s):", "error"
            )
            for key, msg in errors:
                self._batch_log_write(f"  {key}: {msg}", "error")
        else:
            self._batch_log_write(
                f"All {total} plants completed successfully!", "success"
            )

    # ── Theme toggle ──────────────────────────────────────────────────────────
    def _toggle_theme(self):
        ctk.set_appearance_mode(self._theme_switch.get())


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        import argparse
        p = argparse.ArgumentParser()
        p.add_argument("--db",      required=True)
        p.add_argument("--blank",   default="")
        p.add_argument("--folder",  default="")
        p.add_argument("--plant",   default="")
        p.add_argument("--out",     default="")
        p.add_argument("--visible", action="store_true")
        p.add_argument("--all",     action="store_true",
                       help="Run all matched templates in --folder")
        args = p.parse_args()

        if args.all:
            if not args.folder:
                print("--folder is required with --all")
                sys.exit(1)
            matches = find_templates(args.folder)
            if not matches:
                print("No matching _blank.xlsm templates found.")
                sys.exit(1)
            for key, fpath in matches:
                print(f"Filling {key} …")
                fill_plant(args.db, fpath, fpath, REGISTRY[key], visible=args.visible)
                print(f"  → saved: {fpath}")
            print("All done.")
        else:
            key = args.plant.strip().upper()
            if key not in REGISTRY:
                print(f"Unknown plant '{key}'. Available: {sorted(REGISTRY.keys())}")
                sys.exit(1)
            out = args.out.strip() or default_out_path(args.blank)
            fill_plant(args.db, args.blank, out, REGISTRY[key], visible=args.visible)
            print(f"Done. Saved to: {out}")
    else:
        app = App()
        app.mainloop()
