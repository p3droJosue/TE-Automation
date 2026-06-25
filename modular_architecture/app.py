# app.py — TE Template Filler (CustomTkinter UI)
from __future__ import annotations

import json
import os
import sys
import threading

import customtkinter as ctk
from PIL import Image
from tkinter import filedialog

from engine import fill_plant
from plants import REGISTRY, find_templates, match_filename


def _resource_path(rel_path: str) -> str:
    """Resolve an asset path that works in dev and inside a PyInstaller bundle."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

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


# ── Apple-style status palette ────────────────────────────────────────────────
# Soft tinted backgrounds + saturated text. Each color is (light, dark).
_STATUS = {
    "idle":    {"bg": ("#e5e5ea", "#3a3a3c"), "fg": ("#3c3c43", "#ebebf5"), "label": "Ready"},
    "pending": {"bg": ("#e5e5ea", "#3a3a3c"), "fg": ("#6e6e73", "#aeaeb2"), "label": "Pending"},
    "running": {"bg": ("#cce4ff", "#1a3a6e"), "fg": ("#0a84ff", "#64a6ff"), "label": "Running"},
    "done":    {"bg": ("#d1f4d8", "#1b4a2c"), "fg": ("#1b8a3a", "#30d158"), "label": "Done"},
    "failed":  {"bg": ("#ffd6d6", "#5a1d1d"), "fg": ("#cf222e", "#ff453a"), "label": "Failed"},
}

_STATE_ICONS = {"idle": "▢", "pending": "○", "running": "◌", "done": "✓", "failed": "✗"}


def _short_error(msg: str) -> str:
    """Trim a Python exception to a one-line, user-friendly summary."""
    line = (msg or "").strip().splitlines()[0]
    for pref in ("ValueError:", "RuntimeError:", "Exception:",
                 "OSError:", "FileNotFoundError:", "PermissionError:"):
        if line.startswith(pref):
            line = line[len(pref):].strip()
            break
    low = line.lower()
    if "rdp headers missing" in low:
        return "Database is missing required columns."
    if ("another program" in low or "permission denied" in low
            or "being used by" in low or "in use" in low):
        return "Template file is open in Excel — close it and try again."
    if "no such file" in low or "cannot find" in low:
        return "File not found. Check the path."
    return line[:200] if len(line) > 200 else line


# ── Status widgets ────────────────────────────────────────────────────────────
class StatusPill(ctk.CTkLabel):
    """Compact colored pill showing a single state."""

    def __init__(self, master):
        super().__init__(
            master, text=_STATUS["pending"]["label"],
            width=90, height=24,
            corner_radius=12,
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=_STATUS["pending"]["bg"],
            text_color=_STATUS["pending"]["fg"],
        )

    def set_state(self, state: str, text: str | None = None) -> None:
        cfg = _STATUS.get(state, _STATUS["pending"])
        self.configure(
            text=text or cfg["label"],
            fg_color=cfg["bg"],
            text_color=cfg["fg"],
        )


class PlantStatusRow(ctk.CTkFrame):
    """One row in the Run All list: plant name + brief detail + status pill."""

    def __init__(self, master, plant_key: str):
        super().__init__(master, fg_color=("gray95", "gray17"), corner_radius=10)
        self.plant_key = plant_key
        self.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self,
            text=plant_key.replace("_", " ").title(),
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=14, pady=10)

        self._detail = ctk.CTkLabel(
            self, text="",
            text_color=("gray40", "gray65"),
            font=ctk.CTkFont(size=11),
            anchor="w",
        )
        self._detail.grid(row=0, column=1, sticky="w", padx=4, pady=10)

        self._pill = StatusPill(self)
        self._pill.grid(row=0, column=2, padx=14, pady=10, sticky="e")

    def set_state(self, state: str, detail: str = "") -> None:
        self._pill.set_state(state)
        self._detail.configure(text=detail)


class StatusCard(ctk.CTkFrame):
    """Hero status card for the single-plant tab."""

    def __init__(self, master):
        super().__init__(master, corner_radius=14, fg_color=("gray95", "gray17"))
        self.grid_columnconfigure(1, weight=1)

        self._icon = ctk.CTkLabel(
            self, text=_STATE_ICONS["idle"],
            font=ctk.CTkFont(size=40, weight="bold"),
            width=64,
            text_color=_STATUS["idle"]["fg"],
        )
        self._icon.grid(row=0, column=0, rowspan=2, padx=(22, 12), pady=22)

        self._title = ctk.CTkLabel(
            self, text="Ready",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        )
        self._title.grid(row=0, column=1, sticky="ew", pady=(22, 2), padx=(0, 22))

        self._detail = ctk.CTkLabel(
            self, text="Select files and click Run Fill.",
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=12),
            anchor="w",
            wraplength=600,
            justify="left",
        )
        self._detail.grid(row=1, column=1, sticky="ew", pady=(0, 22), padx=(0, 22))

    def set_state(self, state: str, title: str, detail: str = "") -> None:
        cfg = _STATUS.get(state, _STATUS["idle"])
        self._icon.configure(text=_STATE_ICONS.get(state, "▢"), text_color=cfg["fg"])
        self._title.configure(text=title)
        self._detail.configure(text=detail)


# ── Main App ──────────────────────────────────────────────────────────────────
class App(ctk.CTk):
    SIDEBAR_W = 200
    ACCENT    = "#1f6feb"

    def __init__(self):
        super().__init__()
        self.title("TE Template Filler — Mexico")
        # Default size is large enough to show the full sidebar (logo + 15
        # plants + theme button) plus the main panel without scrolling. The
        # user can still maximize or resize; the sidebar plant list scrolls
        # when the window is too short to fit everything.
        self.geometry("1150x820")
        self.minsize(900, 620)
        self.resizable(True, True)

        self._cfg = _load_config()
        self._running = False
        self._selected_plant: str | None = None

        self._batch_rows: dict[str, PlantStatusRow] = {}
        self._batch_matches: list[tuple[str, str]] = []

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
        sb.grid_columnconfigure(0, weight=1)
        # Row 6 (spacer) absorbs extra height so the plant list sits snug
        # under the SELECT PLANT header. The theme button lives in the top
        # section (row 3) so it's visible at any window size — putting it
        # at the bottom of the sidebar made it disappear on short windows.
        sb.grid_rowconfigure(6, weight=1)

        logo_path = _resource_path("assets/pepsico_logo.png")
        if os.path.exists(logo_path):
            self._logo_img = ctk.CTkImage(
                light_image=Image.open(logo_path),
                dark_image=Image.open(logo_path),
                size=(160, 48),
            )
            ctk.CTkLabel(sb, image=self._logo_img, text="").grid(
                row=0, column=0, padx=16, pady=(18, 8), sticky="w"
            )

        ctk.CTkLabel(
            sb, text="TE Filler", font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=1, column=0, padx=16, pady=(4, 4), sticky="w")
        ctk.CTkLabel(
            sb, text="Mexico Sites", font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=2, column=0, padx=16, pady=(0, 12), sticky="w")

        # Theme toggle lives in the top section so the user can always see
        # it, even when the window is short or the plant list overflows.
        self._theme_btn = ctk.CTkButton(
            sb, text="", width=168, height=34, corner_radius=18,
            fg_color=("gray88", "gray22"),
            text_color=("gray10", "gray90"),
            hover_color=("gray80", "gray32"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._toggle_theme,
        )
        self._theme_btn.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="w")
        self._sync_theme_button()

        ctk.CTkLabel(
            sb, text="SELECT PLANT", font=ctk.CTkFont(size=10, weight="bold"),
            text_color="gray"
        ).grid(row=4, column=0, padx=16, pady=(0, 4), sticky="w")

        # Fixed height (~10 plant buttons + scrollbar) keeps the sidebar
        # predictable. sticky="new" + height= prevents vertical stretching.
        plant_list = ctk.CTkScrollableFrame(
            sb, fg_color="transparent", label_text="", height=360,
        )
        plant_list.grid(row=5, column=0, sticky="new", padx=4, pady=0)
        plant_list.grid_columnconfigure(0, weight=1)

        plants = sorted(REGISTRY.keys())
        self._plant_btns: dict[str, ctk.CTkButton] = {}
        for i, key in enumerate(plants):
            btn = ctk.CTkButton(
                plant_list,
                text=key.replace("_", " ").title(),
                anchor="w",
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray80", "gray25"),
                corner_radius=6,
                command=lambda k=key: self._select_plant(k),
            )
            btn.grid(row=i, column=0, padx=4, pady=2, sticky="ew")
            self._plant_btns[key] = btn

    def _sync_theme_button(self):
        """Show the *next* action so the user knows what the click will do."""
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        self._theme_btn.configure(text="☀  Light mode" if is_dark else "☾  Dark mode")

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

        file_row(1, "Database (.xlsm):",   self._db_var,    self._browse_db)
        file_row(2, "Template (.xlsm):",   self._blank_var, self._browse_blank)

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
        self._progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        self._progress.set(0)

        self._status_card = StatusCard(parent)
        self._status_card.grid(row=7, column=0, columnspan=3, sticky="nsew")

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

        self._batch_summary = ctk.CTkLabel(
            parent,
            text="No templates loaded yet.",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        self._batch_summary.grid(row=6, column=0, columnspan=3, sticky="w", padx=4, pady=(2, 8))

        self._batch_list = ctk.CTkScrollableFrame(
            parent, label_text="", fg_color="transparent",
        )
        self._batch_list.grid(row=7, column=0, columnspan=3, sticky="nsew")
        self._batch_list.grid_columnconfigure(0, weight=1)

    # ── Batch row management ──────────────────────────────────────────────────
    def _rebuild_batch_rows(self, matches: list[tuple[str, str]]) -> None:
        for child in self._batch_list.winfo_children():
            child.destroy()
        self._batch_rows = {}
        self._batch_matches = list(matches)

        if not matches:
            self._batch_summary.configure(
                text="No matching templates in folder.",
                text_color=_STATUS["failed"]["fg"],
            )
            return

        self._batch_summary.configure(
            text=f"{len(matches)} template(s) ready",
            text_color=("gray10", "gray90"),
        )

        for i, (key, fpath) in enumerate(matches):
            row = PlantStatusRow(self._batch_list, key)
            row.grid(row=i, column=0, sticky="ew", padx=2, pady=3)
            row.set_state("pending", os.path.basename(fpath))
            self._batch_rows[key] = row

    def _update_batch_summary(self, done: int, failed: int, total: int) -> None:
        pending = total - done - failed
        parts = [f"{done}/{total} done"]
        if failed:
            parts.append(f"{failed} failed")
        if pending:
            parts.append(f"{pending} pending")
        color = _STATUS["failed"]["fg"] if failed else ("gray10", "gray90")
        self._batch_summary.configure(text=" • ".join(parts), text_color=color)

    # ── State / restore ───────────────────────────────────────────────────────
    def _restore_state(self):
        mode = self._cfg.get("appearance_mode", "dark")
        if mode in ("light", "dark"):
            ctk.set_appearance_mode(mode)
            self._sync_theme_button()

        db = self._cfg.get("last_db", "")
        if db and os.path.exists(db):
            self._db_var.set(db)
            self._batch_db_var.set(db)

        folder = self._cfg.get("last_folder", "")
        if folder and os.path.isdir(folder):
            self._batch_folder_var.set(folder)
            self._preview_matches()

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
            self._batch_db_var.set(path)
            self._cfg["last_db"] = path
            _save_config(self._cfg)

    def _browse_blank(self):
        path = filedialog.askopenfilename(
            title="Select TE Template (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._blank_var.set(path)

    # ── File pickers — batch ──────────────────────────────────────────────────
    def _browse_batch_db(self):
        path = filedialog.askopenfilename(
            title="Select Database (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")],
        )
        if path:
            self._batch_db_var.set(path)
            self._db_var.set(path)
            self._cfg["last_db"] = path
            _save_config(self._cfg)

    def _browse_batch_folder(self):
        folder = filedialog.askdirectory(title="Select Templates Folder")
        if folder:
            self._batch_folder_var.set(folder)
            self._cfg["last_folder"] = folder
            _save_config(self._cfg)
            self._preview_matches()

    # ── Preview matches ───────────────────────────────────────────────────────
    def _preview_matches(self):
        folder = self._batch_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            self._batch_summary.configure(text="No folder selected.", text_color="gray")
            self._rebuild_batch_rows([])
            return
        self._rebuild_batch_rows(find_templates(folder))

    # ── Run — single plant ────────────────────────────────────────────────────
    def _run(self):
        if self._running:
            return

        db        = self._db_var.get().strip()
        blank     = self._blank_var.get().strip()
        plant_key = self._selected_plant

        if not db or not os.path.exists(db):
            self._status_card.set_state(
                "failed", "Missing database",
                "Pick a valid .xlsm database file before running.",
            )
            return
        if not blank or not os.path.exists(blank):
            self._status_card.set_state(
                "failed", "Missing template",
                "Pick a valid .xlsm template file before running.",
            )
            return
        if not plant_key or plant_key not in REGISTRY:
            self._status_card.set_state(
                "failed", "No plant selected",
                "Choose a plant from the sidebar first.",
            )
            return

        pretty = plant_key.replace("_", " ").title()
        self._status_card.set_state(
            "running", f"Filling {pretty}…",
            "This usually takes 30–60 seconds. Excel runs in the background.",
        )

        self._running = True
        self._run_btn.configure(state="disabled", text="Running…")
        self._progress.configure(mode="indeterminate")
        self._progress.start()

        def _worker():
            try:
                fill_plant(db, blank, blank, REGISTRY[plant_key],
                           visible=self._visible_var.get())
                self.after(0, self._on_success, blank)
            except Exception as exc:
                self.after(0, self._on_error, str(exc))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_success(self, out_path: str):
        self._progress.stop()
        self._progress.set(1)
        self._run_btn.configure(state="normal", text="Run Fill")
        self._running = False
        pretty = (self._selected_plant or "").replace("_", " ").title()
        self._status_card.set_state(
            "done", f"{pretty} — Done",
            f"Saved {os.path.basename(out_path)}.",
        )

    def _on_error(self, exc_msg: str):
        self._progress.stop()
        self._progress.set(0)
        self._run_btn.configure(state="normal", text="Run Fill")
        self._running = False
        self._status_card.set_state(
            "failed", "Something went wrong",
            _short_error(exc_msg),
        )

    # ── Run All ───────────────────────────────────────────────────────────────
    def _run_all(self):
        if self._running:
            return

        db     = self._batch_db_var.get().strip()
        folder = self._batch_folder_var.get().strip()

        if not db or not os.path.exists(db):
            self._batch_summary.configure(
                text="Missing database — pick a .xlsm file.",
                text_color=_STATUS["failed"]["fg"],
            )
            return
        if not folder or not os.path.isdir(folder):
            self._batch_summary.configure(
                text="Missing templates folder.",
                text_color=_STATUS["failed"]["fg"],
            )
            return

        matches = find_templates(folder)
        if not matches:
            self._batch_summary.configure(
                text="No matching .xlsm templates in folder.",
                text_color=_STATUS["failed"]["fg"],
            )
            return

        self._rebuild_batch_rows(matches)

        self._running = True
        self._run_all_btn.configure(state="disabled", text="Running…")
        self._batch_progress.set(0)

        total = len(matches)

        def _worker():
            done = 0
            failed = 0

            for i, (key, fpath) in enumerate(matches, start=1):
                self.after(0, self._mark_row, key, "running", "Working…")
                try:
                    fill_plant(db, fpath, fpath, REGISTRY[key],
                               visible=self._batch_visible_var.get())
                    done += 1
                    self.after(0, self._mark_row, key, "done",
                               os.path.basename(fpath))
                except Exception as exc:
                    failed += 1
                    self.after(0, self._mark_row, key, "failed",
                               _short_error(str(exc)))
                self.after(0, self._batch_progress.set, i / total)
                self.after(0, self._update_batch_summary, done, failed, total)

            self.after(0, self._on_batch_done, done, failed, total)

        threading.Thread(target=_worker, daemon=True).start()

    def _mark_row(self, key: str, state: str, detail: str = "") -> None:
        row = self._batch_rows.get(key)
        if row is not None:
            row.set_state(state, detail)

    def _on_batch_done(self, done: int, failed: int, total: int) -> None:
        self._running = False
        self._run_all_btn.configure(state="normal", text="Run All")
        self._update_batch_summary(done, failed, total)

    # ── Theme toggle ──────────────────────────────────────────────────────────
    def _toggle_theme(self):
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        new_mode = "light" if is_dark else "dark"
        ctk.set_appearance_mode(new_mode)
        self._sync_theme_button()
        self._cfg["appearance_mode"] = new_mode
        _save_config(self._cfg)


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
