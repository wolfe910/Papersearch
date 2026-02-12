import os
import sqlite3
from pathlib import Path
from PIL import Image
import imagehash

APPDATA = Path(os.getenv("APPDATA")) / "WallpaperFinder"
DB_PATH = APPDATA / "wallpaper_index.db"
TRANSCODED = Path(os.getenv("APPDATA")) / "Microsoft" / "Windows" / "Themes" / "TranscodedWallpaper"


def phash_image(path: Path):
    with Image.open(path) as img:
        img = img.convert("RGB")
        return imagehash.phash(img)


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


def main():
    match, dist = find_best_match()
    if not match:
        print("No match found.")
        return

    source_type, source_path, file_name = match
    print(f"Best match ({dist}):")
    print(f"Type: {source_type}")
    print(f"Source: {source_path}")
    print(f"File: {file_name}")


if __name__ == "__main__":
    main()
