import pathlib
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

from paths import process_paths

PO_ANALYSIS = process_paths.AMAZON_PO_CURRENT_FILE.parent
ARCHIVE = process_paths.AMAZON_PO_DATAWAREHOUSE_ARCHIVE_DIR


def make_archive_name(src_path: pathlib.Path) -> pathlib.Path:
    return ARCHIVE / src_path.name


def unique_path(p: pathlib.Path) -> pathlib.Path:
    if not p.exists():
        return p
    base = p.stem
    suffix = p.suffix
    i = 1
    while True:
        candidate = p.with_name(f"{base}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def main():
    # GUI: pick the new vendor file first
    root = tk.Tk()
    root.withdraw()
    src_file = filedialog.askopenfilename(
        title="Select new Amazon Vendor Central PO Report (CSV)",
        initialdir=str(PO_ANALYSIS) if PO_ANALYSIS.exists() else None,
        filetypes=[("CSV Files", "*.csv"), ("All files", "*.*")],
    )
    if not src_file:
        print("No file selected. Exiting.")
        return

    src_path = pathlib.Path(src_file)
    if not src_path.exists():
        messagebox.showerror("Error", f"Selected file does not exist:\n{src_path}")
        return

    # Ensure analysis folder exists
    PO_ANALYSIS.mkdir(parents=True, exist_ok=True)
    ARCHIVE.mkdir(parents=True, exist_ok=True)

    dest = process_paths.AMAZON_PO_CURRENT_FILE
    archive_target = unique_path(make_archive_name(src_path))
    try:
        shutil.copy2(str(src_path), str(dest))
        shutil.copy2(str(src_path), str(archive_target))
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"Failed to copy new file into output folders:\n{e}",
        )
        return

    messagebox.showinfo(
        "Done",
        (
            "New file copied to:\n"
            f"{dest}\n"
            f"{archive_target}"
        ),
    )
    print("Complete.")


if __name__ == "__main__":
    main()
