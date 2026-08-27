"""
Status Report Tracker Agent — desktop widget
---------------------------------------------
A small floating tray icon. Click it, pick a tab (project/category),
fill in a short form built from that tab's column headers, submit —
it's appended as a row in the shared Google Sheet via the Apps Script
web app.

Run directly:
    python status_widget.py

Config:
    Copy config.example.json to config.json (same folder) and fill in
    WEBHOOK_URL (your deployed Apps Script /exec URL) and TOKEN (must
    match SHARED_SECRET in Code.gs).

Packaging into a standalone app: see build_windows.bat / build_mac.sh.
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

import requests
from PIL import Image, ImageDraw
import pystray

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise SystemExit(
            f"Missing config.json next to this script.\n"
            f"Copy config.example.json to config.json and fill in WEBHOOK_URL / TOKEN.\n"
            f"Expected at: {CONFIG_PATH}"
        )
    with open(CONFIG_PATH, "r") as f:
        return json.load(f)


CONFIG = load_config()
WEBHOOK_URL = CONFIG["WEBHOOK_URL"].rstrip("/")
TOKEN = CONFIG.get("TOKEN", "")
YOUR_NAME = CONFIG.get("YOUR_NAME", "")  # optional: prefill a "Name" field if present


def api_get(action, **params):
    params["action"] = action
    resp = requests.get(WEBHOOK_URL, params=params, timeout=15)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Unknown API error"))
    return data


def api_post(tab, values):
    payload = {"token": TOKEN, "tab": tab, "values": values}
    resp = requests.post(WEBHOOK_URL, data=json.dumps(payload), timeout=15)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Unknown API error"))
    return data


# --------------------------- UI ---------------------------------------

class StatusEntryApp:
    """Owns the single hidden Tk root so tray-triggered popups behave on
    both Windows and macOS (each new Toplevel is parented to this root)."""

    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # no visible main window, just the tray icon

    def run(self):
        self.root.mainloop()

    def show_tab_picker(self):
        self.root.after(0, self._show_tab_picker)

    # -- everything below runs on the Tk main thread via `after()` --

    def _show_tab_picker(self):
        try:
            tabs = api_get("tabs")["tabs"]
        except Exception as exc:
            messagebox.showerror("Status Update", f"Couldn't reach the sheet:\n{exc}")
            return
        if not tabs:
            messagebox.showinfo("Status Update", "No tabs found in the sheet.")
            return

        win = tk.Toplevel(self.root)
        win.title("Status Update — pick a tab")
        win.attributes("-topmost", True)
        tk.Label(win, text="What are you logging status for?", padx=16, pady=10).pack()

        for tab in tabs:
            ttk.Button(
                win, text=tab, width=30,
                command=lambda t=tab, w=win: self._open_form(t, w)
            ).pack(padx=16, pady=4)

        ttk.Button(win, text="Cancel", command=win.destroy).pack(pady=(4, 12))

    def _open_form(self, tab, picker_window):
        try:
            columns = api_get("columns", tab=tab)["columns"]
        except Exception as exc:
            messagebox.showerror("Status Update", f"Couldn't load fields:\n{exc}")
            return
        picker_window.destroy()

        win = tk.Toplevel(self.root)
        win.title(f"Status Update — {tab}")
        win.attributes("-topmost", True)

        tk.Label(win, text=tab, font=("", 12, "bold"), padx=16, pady=(12, 4)).pack()

        entries = {}
        form = tk.Frame(win, padx=16, pady=4)
        form.pack(fill="x")

        for col in columns:
            row = tk.Frame(form)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=col, width=18, anchor="w").pack(side="left")
            var = tk.Text(row, height=1, width=32, wrap="word")
            if col.strip().lower() in ("name", "your name") and YOUR_NAME:
                var.insert("1.0", YOUR_NAME)
            var.pack(side="left", fill="x", expand=True)
            entries[col] = var

        def submit():
            values = {col: widget.get("1.0", "end").strip() for col, widget in entries.items()}
            try:
                api_post(tab, values)
            except Exception as exc:
                messagebox.showerror("Status Update", f"Save failed:\n{exc}")
                return
            win.destroy()
            messagebox.showinfo("Status Update", f"Saved to \"{tab}\". Thanks!")

        btns = tk.Frame(win, padx=16, pady=12)
        btns.pack(fill="x")
        ttk.Button(btns, text="Submit", command=submit).pack(side="right")
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right", padx=(0, 8))


# ------------------------- Tray icon ------------------------------------

def make_icon_image():
    """Simple generated icon so we don't need to ship a binary asset."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((2, 2, size - 2, size - 2), fill=(37, 99, 235, 255))
    draw.text((size / 2 - 6, size / 2 - 10), "S", fill="white")
    return img


def main():
    app = StatusEntryApp()

    def on_submit(icon, item):
        app.show_tab_picker()

    def on_quit(icon, item):
        icon.stop()
        app.root.after(0, app.root.quit)

    menu = pystray.Menu(
        pystray.MenuItem("Submit status update", on_submit, default=True),
        pystray.MenuItem("Quit", on_quit),
    )
    icon = pystray.Icon("status_report_tracker", make_icon_image(), "Status Report Tracker", menu)

    # pystray needs its own thread so Tk can own the main thread (required on macOS).
    threading.Thread(target=icon.run, daemon=True).start()
    app.run()


if __name__ == "__main__":
    main()
