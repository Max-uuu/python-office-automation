"""
Batch rename and organise files.

What it does
------------
- Renames files using a consistent pattern:  <prefix>_<YYYY-MM-DD>_<NNN>.<ext>
  The date comes from the file's last-modified time.
- Optionally sorts files into sub-folders by type (Images, Documents, Spreadsheets, ...)
- Runs as a DRY RUN by default: it only prints what it would do.
  Add --apply to actually rename.
- Every real run writes rename_log.csv so changes can be undone with --undo.

Usage
-----
    python batch_rename.py --folder sample_input --prefix invoice            # preview
    python batch_rename.py --folder sample_input --prefix invoice --apply    # rename
    python batch_rename.py --folder sample_input --by-type --apply           # sort into folders
    python batch_rename.py --folder sample_input --undo                      # revert last run
"""

import argparse
import csv
import sys
from datetime import datetime
from pathlib import Path

TYPE_FOLDERS = {
    "Images": {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"},
    "Documents": {".pdf", ".doc", ".docx", ".txt", ".md"},
    "Spreadsheets": {".xls", ".xlsx", ".csv"},
    "Archives": {".zip", ".rar", ".7z"},
}
LOG_NAME = "rename_log.csv"


def type_folder(ext: str) -> str:
    for folder, exts in TYPE_FOLDERS.items():
        if ext.lower() in exts:
            return folder
    return "Other"


def plan(folder: Path, prefix: str, by_type: bool):
    files = sorted(
        (p for p in folder.iterdir() if p.is_file() and p.name != LOG_NAME),
        key=lambda p: (p.stat().st_mtime, p.name),
    )
    moves, counters, taken = [], {}, set()
    for f in files:
        date = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d")
        target_dir = folder / type_folder(f.suffix) if by_type else folder
        key = (target_dir, date)
        counters[key] = counters.get(key, 0) + 1
        new_name = f"{prefix}_{date}_{counters[key]:03d}{f.suffix.lower()}" if prefix else f.name
        target = target_dir / new_name
        # Never overwrite: bump the counter until the name is free.
        while (target.exists() and target != f) or target in taken:
            counters[key] += 1
            new_name = f"{prefix}_{date}_{counters[key]:03d}{f.suffix.lower()}" if prefix else f"{f.stem}_{counters[key]}{f.suffix}"
            target = target_dir / new_name
        taken.add(target)
        if target != f:
            moves.append((f, target))
    return moves


def apply(folder: Path, moves):
    log_path = folder / LOG_NAME
    with log_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["old_path", "new_path"])
        for src, dst in moves:
            dst.parent.mkdir(exist_ok=True)
            src.rename(dst)
            writer.writerow([src.relative_to(folder), dst.relative_to(folder)])
    print(f"\nDone. {len(moves)} file(s) changed. Log written to {log_path.name}")


def undo(folder: Path):
    log_path = folder / LOG_NAME
    if not log_path.exists():
        sys.exit("No rename_log.csv found, nothing to undo.")
    with log_path.open(encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    for row in reversed(rows):
        src, dst = folder / row["new_path"], folder / row["old_path"]
        if src.exists() and not dst.exists():
            src.rename(dst)
            print(f"  {row['new_path']}  ->  {row['old_path']}")
    for sub in TYPE_FOLDERS.keys() | {"Other"}:
        d = folder / sub
        if d.is_dir() and not any(d.iterdir()):
            d.rmdir()
    log_path.unlink()
    print(f"\nUndo complete. {len(rows)} file(s) restored.")


def main():
    parser = argparse.ArgumentParser(description="Batch rename and organise files.")
    parser.add_argument("--folder", required=True, help="Folder to process")
    parser.add_argument("--prefix", default="", help="New name prefix, e.g. invoice")
    parser.add_argument("--by-type", action="store_true", help="Sort files into sub-folders by type")
    parser.add_argument("--apply", action="store_true", help="Actually rename (default is preview only)")
    parser.add_argument("--undo", action="store_true", help="Revert the last applied run")
    args = parser.parse_args()

    folder = Path(args.folder)
    if not folder.is_dir():
        sys.exit(f"Folder not found: {folder}")

    if args.undo:
        undo(folder)
        return
    if not args.prefix and not args.by_type:
        sys.exit("Nothing to do: give --prefix and/or --by-type.")

    moves = plan(folder, args.prefix, args.by_type)
    if not moves:
        print("Nothing to change.")
        return

    print("DRY RUN (no files changed):" if not args.apply else "Renaming:")
    for src, dst in moves:
        print(f"  {src.relative_to(folder)}  ->  {dst.relative_to(folder)}")

    if args.apply:
        apply(folder, moves)
    else:
        print(f"\n{len(moves)} file(s) would change. Add --apply to run it.")


if __name__ == "__main__":
    main()
