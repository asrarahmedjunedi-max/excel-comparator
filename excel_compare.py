"""
Excel Comparison Tool
=====================
Created by Asrar Ahmed Junedi

Compares EVERY worksheet of two Excel workbooks cell-by-cell (matched by
sheet name) and produces a new file, "Comparison_Result.xlsx", based on
File 2's data, with differences highlighted:

  - Changed cell               -> Yellow fill
  - Column header (row 1)      -> Orange fill  (if that column has >=1 diff)
  - Row header (column A)      -> Light red fill (if that row has >=1 diff)

Behavior across sheets:
  - Sheets present in both files are compared cell-by-cell.
  - Sheets that exist only in File 2 are copied through unchanged.
  - Sheets that exist only in File 1 are appended to the output (suffixed
    "(only in File 1)") with every non-empty cell flagged as a difference,
    so nothing from File 1 is silently dropped.

Every output workbook also gets:
  - A locked/protected credit cell in ROW 1 of every sheet (a couple of
    columns to the right of the data, so it's visible immediately and
    doesn't overwrite anything) reading "Created by Asrar Ahmed Junedi"
    (the rest of each sheet stays fully editable -- only that one
    credit cell is locked).
  - A dedicated, fully protected "About" sheet with the same credit.

Works even if matching sheets have different numbers of rows/columns
(missing cells are treated as blank / None and compared accordingly).

Requirements: Python 3.8+, openpyxl, tkinter (bundled with standard Python
on Windows).

Install dependency:
    pip install openpyxl

Run:
    python excel_compare.py

Build EXE using PyInstaller (see build_exe.bat / installer.iss provided
alongside this file):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name "ExcelComparator" excel_compare.py
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.protection import Protection
from openpyxl.utils import get_column_letter


# ----------------------------------------------------------------------
# Fill colors
# ----------------------------------------------------------------------
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")   # changed cells
ORANGE_FILL = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")   # column headers
LIGHT_RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # row headers

CREDIT_TEXT = "Created by Asrar Ahmed Junedi"
SHEET_PROTECT_PASSWORD = "asrar143"  # light-weight lock; change/remove as you like


# ----------------------------------------------------------------------
# Credit / locking helpers
# ----------------------------------------------------------------------
def _add_locked_credit_to_sheet(ws):
    """
    Adds the credit line at the TOP of the sheet -- in row 1, a couple
    of columns to the right of the existing data -- and locks ONLY that
    cell. Placing it in row 1 (rather than below the data) means it's
    visible the moment the sheet is opened, with no scrolling needed.

    Every other existing cell on the sheet is explicitly unlocked first,
    so protection doesn't stop the user from editing/using the rest of
    their comparison data.
    """
    credit_col = (ws.max_column or 0) + 2  # 1 blank column of separation
    credit_cell = ws.cell(row=1, column=credit_col)

    # Unlock every pre-existing cell so sheet protection only bites on
    # the credit cell we're about to lock.
    for row in ws.iter_rows():
        for cell in row:
            if cell is not credit_cell:
                cell.protection = Protection(locked=False)

    credit_cell.value = CREDIT_TEXT
    credit_cell.font = Font(bold=True, italic=True, size=10, color="C00000")
    credit_cell.protection = Protection(locked=True)

    # Widen the column a bit so the text isn't clipped.
    col_letter = get_column_letter(credit_col)
    current_width = ws.column_dimensions[col_letter].width
    needed_width = len(CREDIT_TEXT) * 1.1
    if not current_width or current_width < needed_width:
        ws.column_dimensions[col_letter].width = needed_width

    ws.protection.sheet = True
    ws.protection.password = SHEET_PROTECT_PASSWORD
    ws.protection.enable()


def _add_about_sheet(wb):
    """Adds (or refreshes) a fully-locked 'About' sheet with the credit."""
    if "About" in wb.sheetnames:
        del wb["About"]
    ws = wb.create_sheet("About")

    ws["A1"] = "Excel Comparison Tool"
    ws["A1"].font = Font(bold=True, size=14)

    ws["A2"] = CREDIT_TEXT
    ws["A2"].font = Font(italic=True, size=11, color="404040")

    ws.column_dimensions["A"].width = 40

    for row in ws.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=True)

    ws.protection.sheet = True
    ws.protection.password = SHEET_PROTECT_PASSWORD
    ws.protection.enable()


def _finalize_credit(out_wb):
    """Applies the locked credit line to every data sheet, then adds
    the dedicated About sheet last."""
    for ws in list(out_wb.worksheets):
        _add_locked_credit_to_sheet(ws)
    _add_about_sheet(out_wb)


# ----------------------------------------------------------------------
# Core comparison logic
# ----------------------------------------------------------------------
def _compare_sheet(ws1, ws2, out_ws):
    """
    Compares two worksheets (ws1, ws2 -- either may be None if a sheet is
    missing from one workbook) cell by cell, writing results/highlights
    into out_ws (built from File 2's data). Returns the diff count for
    this sheet.
    """
    max_row = max(ws1.max_row if ws1 else 0, ws2.max_row if ws2 else 0)
    max_col = max(ws1.max_column if ws1 else 0, ws2.max_column if ws2 else 0)

    rows_with_diff = set()
    cols_with_diff = set()
    diff_count = 0

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v1 = ws1.cell(row=r, column=c).value if ws1 else None
            v2 = ws2.cell(row=r, column=c).value if ws2 else None

            if v1 != v2:
                diff_count += 1
                rows_with_diff.add(r)
                cols_with_diff.add(c)

                out_cell = out_ws.cell(row=r, column=c)
                out_cell.value = f"{v1} -> {v2}"
                out_cell.fill = YELLOW_FILL

    # Highlight column headers (row 1) for every column that had a diff
    for c in cols_with_diff:
        out_ws.cell(row=1, column=c).fill = ORANGE_FILL

    # Highlight row headers (column A) for every row that had a diff
    for r in rows_with_diff:
        out_ws.cell(row=r, column=1).fill = LIGHT_RED_FILL

    return diff_count


def compare_excel_files(file1_path, file2_path, output_path="Comparison_Result.xlsx"):
    """
    Compares EVERY worksheet of file1 against the corresponding worksheet
    (matched by name) in file2, cell by cell.

    - Sheets present in both files: compared cell-by-cell as before.
    - Sheets present only in File 2: copied through unchanged (nothing to
      compare against, so no highlighting).
    - Sheets present only in File 1: added to the output (empty in File 2)
      and every non-empty File 1 cell is reported as removed, so the
      difference is not silently lost.

    Produces output_path, a workbook based on File 2's structure/values,
    with differences highlighted per-sheet, PLUS a locked credit line on
    every sheet and a protected "About" sheet.

    Returns: (output_path, total_diff_count)
    """
    wb1 = load_workbook(file1_path, data_only=True)
    wb2 = load_workbook(file2_path, data_only=True)

    # Output workbook is built from File 2 (values only) so we can freely
    # format it. Sheet order follows File 2, then any File-1-only sheets
    # are appended at the end.
    out_wb = load_workbook(file2_path, data_only=True)

    total_diff_count = 0

    # --- 1) Sheets that exist in File 2 (compare against File 1 if present) ---
    for sheet_name in wb2.sheetnames:
        ws2 = wb2[sheet_name]
        ws1 = wb1[sheet_name] if sheet_name in wb1.sheetnames else None
        out_ws = out_wb[sheet_name]

        if ws1 is not None:
            total_diff_count += _compare_sheet(ws1, ws2, out_ws)
        # else: sheet only in File 2 -> left as-is, no highlighting.

    # --- 2) Sheets that exist ONLY in File 1 -> append to output ---
    for sheet_name in wb1.sheetnames:
        if sheet_name in wb2.sheetnames:
            continue  # already handled above

        ws1 = wb1[sheet_name]
        out_ws = out_wb.create_sheet(title=f"{sheet_name} (only in File 1)")

        # Copy File 1's values in, then mark every non-empty cell as a
        # difference (since File 2 has nothing here).
        diff_count = 0
        for r in range(1, ws1.max_row + 1):
            for c in range(1, ws1.max_column + 1):
                v1 = ws1.cell(row=r, column=c).value
                if v1 is not None:
                    out_cell = out_ws.cell(row=r, column=c)
                    out_cell.value = f"{v1} -> (missing)"
                    out_cell.fill = YELLOW_FILL
                    diff_count += 1

        if diff_count:
            for c in range(1, ws1.max_column + 1):
                out_ws.cell(row=1, column=c).fill = ORANGE_FILL
            for r in range(1, ws1.max_row + 1):
                out_ws.cell(row=r, column=1).fill = LIGHT_RED_FILL

        total_diff_count += diff_count

    # --- 3) Stamp the locked credit line on every sheet + About sheet ---
    _finalize_credit(out_wb)

    out_wb.save(output_path)
    return output_path, total_diff_count


# ----------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------
class ExcelCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Comparator — by Asrar Ahmed Junedi")
        self.root.geometry("600x290")
        self.root.resizable(False, False)

        self.file1_path = tk.StringVar(value="No file selected")
        self.file2_path = tk.StringVar(value="No file selected")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        title_label = tk.Label(
            self.root, text="Excel File Comparator", font=("Segoe UI", 16, "bold")
        )
        title_label.pack(pady=(15, 5))

        subtitle = tk.Label(
            self.root,
            text="Select two Excel files and compare all worksheets (entire workbook).",
            font=("Segoe UI", 9),
            fg="gray",
        )
        subtitle.pack(pady=(0, 10))

        # --- File 1 row ---
        frame1 = tk.Frame(self.root)
        frame1.pack(fill="x", **pad)
        tk.Button(
            frame1, text="Select File 1", width=15, command=self.select_file1
        ).pack(side="left")
        tk.Label(
            frame1, textvariable=self.file1_path, anchor="w", fg="blue",
            wraplength=380, justify="left"
        ).pack(side="left", padx=10)

        # --- File 2 row ---
        frame2 = tk.Frame(self.root)
        frame2.pack(fill="x", **pad)
        tk.Button(
            frame2, text="Select File 2", width=15, command=self.select_file2
        ).pack(side="left")
        tk.Label(
            frame2, textvariable=self.file2_path, anchor="w", fg="blue",
            wraplength=380, justify="left"
        ).pack(side="left", padx=10)

        # --- Compare button ---
        compare_frame = tk.Frame(self.root)
        compare_frame.pack(pady=20)
        tk.Button(
            compare_frame,
            text="Compare",
            width=20,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("Segoe UI", 10, "bold"),
            command=self.run_comparison,
        ).pack()

        # --- Status label ---
        self.status_var = tk.StringVar(value="")
        tk.Label(
            self.root, textvariable=self.status_var, fg="darkgreen",
            font=("Segoe UI", 9, "italic")
        ).pack(pady=(0, 5))

        # --- Credit footer ---
        tk.Label(
            self.root, text=CREDIT_TEXT,
            fg="gray", font=("Segoe UI", 8)
        ).pack(pady=(0, 10))

    def select_file1(self):
        path = filedialog.askopenfilename(
            title="Select File 1",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file1_path.set(path)

    def select_file2(self):
        path = filedialog.askopenfilename(
            title="Select File 2",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if path:
            self.file2_path.set(path)

    def run_comparison(self):
        f1 = self.file1_path.get()
        f2 = self.file2_path.get()

        if not os.path.isfile(f1) or not os.path.isfile(f2):
            messagebox.showerror(
                "Missing Files", "Please select both File 1 and File 2 before comparing."
            )
            return

        try:
            self.status_var.set("Comparing... please wait.")
            self.root.update_idletasks()

            # Save result next to File 2 by default (output is based on File 2)
            out_dir = os.path.dirname(f2) or os.getcwd()
            output_path = os.path.join(out_dir, "Comparison_Result.xlsx")

            output_path, diff_count = compare_excel_files(f1, f2, output_path)

            self.status_var.set(f"Done. {diff_count} difference(s) found.")
            messagebox.showinfo(
                "Comparison Complete",
                f"Comparison finished.\n\n"
                f"Differences found: {diff_count}\n"
                f"Result saved to:\n{output_path}",
            )
        except Exception as e:
            self.status_var.set("Error during comparison.")
            messagebox.showerror(
                "Error", f"An error occurred while comparing files:\n\n{e}"
            )
            traceback.print_exc()


def main():
    root = tk.Tk()
    app = ExcelCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
