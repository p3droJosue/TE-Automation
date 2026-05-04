# main.py (RUNNER)
from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from engine import fill_plant, default_out_path
from plants import REGISTRY


def run_gui():
    root = tk.Tk()
    root.title("TE Template Filler (Mexico)")
    root.geometry("900x360")

    db_var = tk.StringVar()
    blank_var = tk.StringVar()
    out_var = tk.StringVar()
    plant_var = tk.StringVar(value=sorted(REGISTRY.keys())[0])
    visible_var = tk.BooleanVar(value=False)

    def browse_db():
        path = filedialog.askopenfilename(
            title="Select DB (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")]
        )
        if path:
            db_var.set(path)

    def browse_blank():
        path = filedialog.askopenfilename(
            title="Select Blank TE Template (.xlsm)",
            filetypes=[("Excel Macro-Enabled", "*.xlsm"), ("All files", "*.*")]
        )
        if path:
            blank_var.set(path)
            out_var.set(default_out_path(path))

    def browse_out():
        path = filedialog.asksaveasfilename(
            title="Save Filled TE as...",
            defaultextension=".xlsm",
            filetypes=[("Excel Macro-Enabled", "*.xlsm")]
        )
        if path:
            out_var.set(path)

    def run():
        db = db_var.get().strip()
        blank = blank_var.get().strip()
        out = out_var.get().strip()
        plant_key = plant_var.get().strip().upper()

        if not db or not os.path.exists(db):
            messagebox.showerror("Missing file", "Please select a valid DB .xlsm")
            return
        if not blank or not os.path.exists(blank):
            messagebox.showerror("Missing file", "Please select a valid Blank TE template .xlsm")
            return
        if plant_key not in REGISTRY:
            messagebox.showerror("Plant not found", f"Unknown plant {plant_key}. Available: {sorted(REGISTRY.keys())}")
            return
        if not out:
            out = default_out_path(blank)
            out_var.set(out)

        try:
            btn_run.config(state="disabled")
            root.update_idletasks()
            fill_plant(db, blank, out, REGISTRY[plant_key], visible=visible_var.get())
            messagebox.showinfo("Success", f"Filled TE saved:\n{out}")
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            btn_run.config(state="normal")

    pad = {"padx": 10, "pady": 6}
    frm = ttk.Frame(root)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="DB (.xlsm):").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=db_var, width=95).grid(row=0, column=1, **pad)
    ttk.Button(frm, text="Browse", command=browse_db).grid(row=0, column=2, **pad)

    ttk.Label(frm, text="Blank TE Template (.xlsm):").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=blank_var, width=95).grid(row=1, column=1, **pad)
    ttk.Button(frm, text="Browse", command=browse_blank).grid(row=1, column=2, **pad)

    ttk.Label(frm, text="Output (Filled TE):").grid(row=2, column=0, sticky="w", **pad)
    ttk.Entry(frm, textvariable=out_var, width=95).grid(row=2, column=1, **pad)
    ttk.Button(frm, text="Save as", command=browse_out).grid(row=2, column=2, **pad)

    ttk.Label(frm, text="Plant:").grid(row=3, column=0, sticky="w", **pad)
    ttk.Combobox(frm, textvariable=plant_var, values=sorted(REGISTRY.keys()),
                 state="readonly", width=30).grid(row=3, column=1, sticky="w", **pad)

    ttk.Checkbutton(frm, text="Show Excel while running", variable=visible_var).grid(row=4, column=1, sticky="w", **pad)

    btn_run = ttk.Button(frm, text="Run Fill", command=run)
    btn_run.grid(row=5, column=1, sticky="e", **pad)

    root.mainloop()


def run_cli():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--db", required=True)
    p.add_argument("--blank", required=True)
    p.add_argument("--plant", required=True)
    p.add_argument("--out", default="")
    p.add_argument("--visible", action="store_true")
    args = p.parse_args()

    plant_key = args.plant.strip().upper()
    if plant_key not in REGISTRY:
        raise ValueError(f"Unknown plant '{plant_key}'. Available: {sorted(REGISTRY.keys())}")

    out = args.out.strip() or default_out_path(args.blank)
    fill_plant(args.db, args.blank, out, REGISTRY[plant_key], visible=args.visible)
    print(f"✅ Saved filled TE: {out}")


if __name__ == "__main__":
    if len(sys.argv) == 1:
        run_gui()
    else:
        run_cli()