import os
import sys
import json
import sqlite3
import threading
import subprocess
from datetime import datetime
from pathlib import Path
from zipfile import ZipFile, is_zipfile

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, Image as PILImage
import imagehash
import pystray
from pystray import MenuItem as Item

# ---- PATHS ----
APPDATA = Path(os.getenv("APPDATA")) / "WallpaperFinder"
SETTINGS_PATH = APPDATA / "settings.json"
DB_PATH = APPDATA / "wallpaper_index.db"
TRANSCODED = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Themes" / "TranscodedWallpaper"

ICON_PATH = Path(__file__).parent / "papersearch.ico"
WINDOW_SIZE = "700x500"
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}


# ---- UTIL ----

def ensure_appdata():
    APPDATA.mkdir(parents=True, exist_ok=True)


def load_settings():
    ensure_appdata()
    if SETTINGS_PATH.exists():
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_settings(data):
    ensure_appdata()
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def init_db():
    ensure_appdata()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_type TEXT NOT NULL,
            source_path TEXT NOT NULL,
            file_name TEXT NOT NULL,
            phash TEXT NOT NULL
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_phash ON images(phash)")
    conn.commit()
    conn.close()


def phash_image(path: Path):
    with Image.open(path) as img:
        img = img.convert("RGB")
        return imagehash.phash(img)


def phash_image_from_bytes(data: bytes):
    from io import BytesIO
    with Image.open(BytesIO(data)) as img:
        img = img.convert("RGB")
        return imagehash.phash(img)


def open_in_explorer(full_path: Path):
    subprocess.Popen(["explorer.exe", f"/select,{full_path}"])


def index_source(source_type: str, source_path: Path, log_func=None):
    init_db()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    if log_func:
        log_func(f"Indexing {source_type}: {source_path}")

    if source_type == "zip":
        with ZipFile(source_path, "r") as zf:
            for info in zf.infolist():
                name = info.filename
                ext = Path(name).suffix.lower()
                if ext not in IMAGE_EXTS:
                    continue
                try:
                    data = zf.read(info)
                    h = phash_image_from_bytes(data)
                    cur.execute(
                        "INSERT INTO images (source_type, source_path, file_name, phash) VALUES (?, ?, ?, ?)",
                        ("zip", str(source_path), name, str(h))
                    )
                except Exception:
                    continue
    else:
        for path in source_path.rglob("*"):
            if not path.is_file():
                continue
            ext = path.suffix.lower()
            if ext not in IMAGE_EXTS:
                continue
            try:
                h = phash_image(path)
                rel = path.relative_to(source_path)
                cur.execute(
                    "INSERT INTO images (source_type, source_path, file_name, phash) VALUES (?, ?, ?, ?)",
                    ("folder", str(source_path), str(rel), str(h))
                )
            except Exception:
                continue

    conn.commit()
    conn.close()
    if log_func:
        log_func("Indexing complete.")


def find_best_match():
    if not TRANSCODED.exists():
        raise FileNotFoundError(f"TranscodedWallpaper not found: {TRANSCODED}")
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found: {DB_PATH}")

    wall_hash = phash_image(TRANSCODED)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT source_type, source_path, file_name, phash FROM images")
    rows = cur.fetchall()
    conn.close()

    best = None
    best_dist = 10**9

    for source_type, source_path, file_name, phash_hex in rows:
        try:
            stored_hash = imagehash.hex_to_hash(phash_hex)
            dist = wall_hash - stored_hash
        except Exception:
            continue

        if dist < best_dist:
            best_dist = dist
            best = (source_type, Path(source_path), file_name)

    return best, best_dist


# ---- GUI ----

class WallpaperGUI:
    def __init__(self, root, tray_app):
        self.root = root
        self.tray_app = tray_app

        self.settings = load_settings()
        self.source_type = self.settings.get("source_type")
        self.source_path = Path(self.settings["source_path"]) if "source_path" in self.settings else None

        self.root.title("Wallpaper Source Finder")
        self.root.geometry(WINDOW_SIZE)
        self.root.resizable(False, False)

        ttk.Label(root, text="Source (ZIP or Folder):").pack(anchor="w", padx=10, pady=(10, 0))
        self.source_entry = ttk.Entry(root, width=95)
        self.source_entry.pack(padx=10)

        ttk.Button(root, text="Change Source", command=self.change_source).pack(pady=3)

        ttk.Label(root, text="Directory Path:").pack(anchor="w", padx=10, pady=(10, 0))
        self.dir_entry = ttk.Entry(root, width=95)
        self.dir_entry.pack(padx=10)

        ttk.Button(root, text="Copy Directory Path",
                   command=lambda: self.copy_to_clipboard(self.dir_entry.get())).pack(pady=3)

        ttk.Label(root, text="Filename:").pack(anchor="w", padx=10, pady=(10, 0))
        self.file_entry = ttk.Entry(root, width=95)
        self.file_entry.pack(padx=10)

        ttk.Button(root, text="Copy Filename",
                   command=lambda: self.copy_to_clipboard(self.file_entry.get())).pack(pady=3)

        ttk.Button(root, text="Refresh Now", command=self.refresh).pack(pady=10)

        self.auto_var = tk.BooleanVar()
        self.auto_var.set(self.settings.get("auto_refresh", False))
        ttk.Checkbutton(root, text="Enable Auto-Refresh", variable=self.auto_var,
                        command=self.toggle_auto).pack()

        ttk.Label(root, text="Refresh Interval (seconds):").pack()
        self.interval_entry = ttk.Entry(root, width=10)
        self.interval_entry.insert(0, str(self.settings.get("interval", 5)))
        self.interval_entry.pack()

        ttk.Label(root, text="Console Log:").pack(anchor="w", padx=10, pady=(10, 0))
        self.console = tk.Text(root, height=10, width=95, state="disabled")
        self.console.pack(padx=10, pady=5)

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        if not self.source_path:
            self.first_run_setup()
        else:
            self.source_entry.insert(0, f"{self.source_type}: {self.source_path}")

        if self.auto_var.get():
            self.schedule_refresh()
        else:
            self.refresh()

    def log(self, text):
        timestamp = datetime.now().strftime("%H:%M:%S")
        message = f"[{timestamp}] {text}\n"
        self.console.config(state="normal")
        self.console.insert(tk.END, message)
        self.console.see(tk.END)
        self.console.config(state="disabled")

    def copy_to_clipboard(self, text):
        if not text:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(str(text))
        self.log(f"Copied to clipboard: {text}")

    def first_run_setup(self):
        self.log("First run: selecting image source.")
        messagebox.showinfo(
            "Select Image Source",
            "Select a ZIP file or a folder containing your wallpapers."
        )

        path_str = filedialog.askopenfilename(
            title="Select ZIP file (or cancel to choose folder)",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )

        if path_str and is_zipfile(path_str):
            self.source_type = "zip"
            self.source_path = Path(path_str)
        else:
            folder_str = filedialog.askdirectory(
                title="Select folder containing images"
            )
            if not folder_str:
                messagebox.showerror("Error", "No source selected. Exiting.")
                self.root.after(100, self.root.quit)
                return
            self.source_type = "folder"
            self.source_path = Path(folder_str)

        self.settings["source_type"] = self.source_type
        self.settings["source_path"] = str(self.source_path)
        save_settings(self.settings)

        self.source_entry.delete(0, tk.END)
        self.source_entry.insert(0, f"{self.source_type}: {self.source_path}")

        threading.Thread(
            target=index_source,
            args=(self.source_type, self.source_path, self.log),
            daemon=True
        ).start()

    def change_source(self):
        messagebox.showinfo(
            "Change Image Source",
            "Select a new ZIP file or folder. This will rebuild the index."
        )

        path_str = filedialog.askopenfilename(
            title="Select ZIP file (or cancel to choose folder)",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )

        new_type = None
        new_path = None

        if path_str and is_zipfile(path_str):
            new_type = "zip"
            new_path = Path(path_str)
        else:
            folder_str = filedialog.askdirectory(
                title="Select folder containing images"
            )
            if not folder_str:
                self.log("Change source cancelled.")
                return
            new_type = "folder"
            new_path = Path(folder_str)

        self.source_type = new_type
        self.source_path = new_path
        self.settings["source_type"] = self.source_type
        self.settings["source_path"] = str(self.source_path)
        save_settings(self.settings)

        self.source_entry.delete(0, tk.END)
        self.source_entry.insert(0, f"{self.source_type}: {self.source_path}")

        if DB_PATH.exists():
            DB_PATH.unlink()
            self.log("Old index removed.")

        threading.Thread(
            target=index_source,
            args=(self.source_type, self.source_path, self.log),
            daemon=True
        ).start()

    def refresh(self):
        self.log("Refreshing...")
        try:
            match, dist = find_best_match()
        except FileNotFoundError as e:
            self.log(f"ERROR: {e}")
            return
        except Exception as e:
            self.log(f"ERROR: {e}")
            return

        if not match:
            self.log("No match found.")
            return

        source_type, source_path, file_name = match
        self.log(f"Match: {source_type} | {source_path} | {file_name} (dist {dist})")

        if source_type == "zip":
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, str(source_path))
        else:
            base_dir = source_path
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, str(base_dir))

        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_name)

        self.log("Refresh complete.")

    def toggle_auto(self):
        self.settings["auto_refresh"] = bool(self.auto_var.get())
        try:
            self.settings["interval"] = int(self.interval_entry.get())
        except ValueError:
            self.settings["interval"] = 5
        save_settings(self.settings)

        if self.auto_var.get():
            self.log("Auto-refresh ENABLED.")
            self.schedule_refresh()
        else:
            self.log("Auto-refresh DISABLED.")

    def schedule_refresh(self):
        if not self.auto_var.get():
            return
        try:
            interval = int(self.interval_entry.get()) * 1000
        except ValueError:
            interval = 5000
            self.log("Invalid interval. Defaulting to 5 seconds.")
        self.refresh()
        self.root.after(interval, self.schedule_refresh)

    def quick_locate_and_copy(self):
        self.refresh()
        filename = self.file_entry.get().strip()
        if filename:
            self.copy_to_clipboard(filename)

        if not self.source_type or not self.source_path:
            self.log("No source configured.")
            return

        if self.source_type == "zip":
            open_in_explorer(self.source_path)
            self.log("Quick locate: ZIP highlighted.")
        else:
            full_path = self.source_path / filename
            open_in_explorer(full_path)
            self.log("Quick locate: file highlighted.")

    def on_close(self):
        self.root.withdraw()
        self.log("Window hidden to tray.")


# ---- TRAY APP ----

class TrayApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        self.gui = None

        try:
            img = PILImage.open(ICON_PATH)
        except Exception:
            img = PILImage.new("RGB", (64, 64), color=(255, 255, 255))

        self.icon = pystray.Icon(
            "WallpaperFinder",
            img,
            "Wallpaper Finder",
            menu=pystray.Menu(
                Item("Show Window", self.show_window),
                Item("Quick Locate & Copy", self.quick_locate_and_copy),
                Item("Refresh Now", self.refresh_now),
                Item("Exit", self.exit_app)
            )
        )

        threading.Thread(target=self.icon.run, daemon=True).start()
        self.root.after(200, self.ensure_gui)

    def ensure_gui(self):
        if self.gui is None:
            self.gui = WallpaperGUI(self.root, self)

    def show_window(self, icon, item):
        if self.gui:
            self.gui.root.deiconify()
            self.gui.root.lift()
            self.gui.root.focus_force()
            self.gui.log("Window shown from tray.")

    def refresh_now(self, icon, item):
        if self.gui:
            self.gui.refresh()

    def quick_locate_and_copy(self, icon, item):
        if self.gui:
            self.gui.quick_locate_and_copy()

    def exit_app(self, icon, item):
        if self.gui:
            self.gui.log("Exiting application.")
        self.icon.stop()
        self.root.quit()
        sys.exit(0)


if __name__ == "__main__":
    app = TrayApp()
    app.root.mainloop()
