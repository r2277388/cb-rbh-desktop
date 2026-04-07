import datetime
import pathlib
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

from paths import process_paths

PO_ANALYSIS = process_paths.AMAZON_PO_CURRENT_PREORDERS_FILE.parent
PO_ANALYSIS_MIRROR = process_paths.AMAZON_PO_DATAWAREHOUSE_ANALYSIS_FILE.parent
ARCHIVE = PO_ANALYSIS / "po_archive"
ARCHIVE_MIRROR = PO_ANALYSIS_MIRROR / "po_archive"
PRIOR_GLOB = (
    "current_amaz_preorders.*"  # matches current_amaz_preorders.csv, .xlsx, etc.
)


def make_archive_name(prior_path: pathlib.Path) -> pathlib.Path:
    mtime = prior_path.stat().st_mtime
    date_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y%m%d")
    return ARCHIVE / f"PurchaseOrderItems_{date_str}{prior_path.suffix}"


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
        title="Select new 'current_amaz_preorders' file",
        initialdir=str(PO_ANALYSIS) if PO_ANALYSIS.exists() else None,
        filetypes=[("All files", "*.*")],
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
    PO_ANALYSIS_MIRROR.mkdir(parents=True, exist_ok=True)
    ARCHIVE.mkdir(parents=True, exist_ok=True)
    ARCHIVE_MIRROR.mkdir(parents=True, exist_ok=True)

    # Find prior current_amaz_preorders file (if any)
    priors = list(PO_ANALYSIS.glob(PRIOR_GLOB))
    if priors:
        # if multiple, choose the most recently modified
        prior = max(priors, key=lambda p: p.stat().st_mtime)
        archive_target = make_archive_name(prior)
        archive_target = unique_path(archive_target)
        mirror_archive_target = ARCHIVE_MIRROR / archive_target.name
        mirror_archive_target = unique_path(mirror_archive_target)
        try:
            shutil.copy2(str(prior), str(mirror_archive_target))
            shutil.move(str(prior), str(archive_target))
            print(
                f"Moved prior file {prior.name} -> {archive_target}"
                f"\nCopied archive mirror -> {mirror_archive_target}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to archive prior file:\n{e}")
            return
    else:
        print("No prior current_amaz_preorders file found; continuing.")
        archive_target = None
        mirror_archive_target = None

    dest = PO_ANALYSIS / f"current_amaz_preorders{src_path.suffix}"
    mirror_dest = PO_ANALYSIS_MIRROR / f"current_amaz_preorders{src_path.suffix}"
    try:
        shutil.copy2(str(src_path), str(dest))
        shutil.copy2(str(src_path), str(mirror_dest))
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
            f"{mirror_dest}\n\n"
            "Archived prior file to:\n"
            f"{archive_target or '(no prior file found)'}\n"
            f"{mirror_archive_target or '(no prior file found)'}"
        ),
    )
    print("Complete.")


if __name__ == "__main__":
    main()
