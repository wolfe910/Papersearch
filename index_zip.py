import os
import sqlite3
from pathlib import Path
from zipfile import ZipFile, is_zipfile
from PIL import Image
import imagehash
import tkinter as tk
from tkinter import filedialog, messagebox

APPDATA = Path(os.getenv("APPDATA")) / "WallpaperFinder"
DB_PATH = APPDATA / "wallpaper_index.db"
SETTINGS_PATH = APPDATA / "settings.json"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}


def ensure_appdata():
    APPDATA.mkdir(parents=True, exist_ok=True)


def init_db():
    ensure_appdata()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_type TEXT NOT NULL,   -- 'zip' or 'folder'
            source_path TEXT NOT NULL,   -- full path to zip or folder
            file_name TEXT NOT NULL,     -- internal path or filename
            phash TEXT NOT NULL
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_phash ON images(phash)")
    conn.commit()
    conn.close()


def phash_image_from_bytes(data):
    from io import BytesIO
    with Image.open(BytesIO(data)) as img:
        img = img.convert("RGB")
        return imagehash.phash(img)


def phash_image_from_path(path: Path):
    with Image.open(path) as img:
        img = img.convert("RGB")
        return imagehash.phash(img)


def index_zip(zip_path: Path):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    with ZipFile(zip_path, "r") as zf:
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
                    ("zip", str(zip_path), name, str(h))
                )
            except Exception:
                continue

    conn.commit()
    conn.close()


def index_folder(folder: Path):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    for path in folder.rglob("*"):
        if not path.is_file():
            continue
        ext = path.suffix.lower()
        if ext not in IMAGE_EXTS:
            continue
        try:
            h = phash_image_from_path(path)
            rel = path.relative_to(folder)
            cur.execute(
                "INSERT INTO images (source_type, source_path, file_name, phash) VALUES (?, ?, ?, ?)",
                ("folder", str(folder), str(rel), str(h))
            )
        except Exception:
            continue

    conn.commit()
    conn.close()


def main():
    ensure_appdata()
    init_db()

    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Select Image Source",
        "Select a ZIP file or a folder containing your wallpapers."
    )

    path_str = filedialog.askopenfilename(
        title="Select ZIP file (or cancel to choose folder)",
        filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
    )

    source_path = None
    if path_str and is_zipfile(path_str):
        source_path = Path(path_str)
        index_zip(source_path)
    else:
        folder_str = filedialog.askdirectory(
            title="Select folder containing images"
        )
        if not folder_str:
            messagebox.showerror("Error", "No source selected.")
            return
        source_path = Path(folder_str)
        index_folder(source_path)

    messagebox.showinfo("Done", f"Indexing complete.\nSource: {source_path}\nDB: {DB_PATH}")


if __name__ == "__main__":
    main()
