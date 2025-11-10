from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Tuple

from .asy_generator import generate_asy_file
from .notebook_generator import generate_liquid_handler_notebook


def show_file_dialog() -> Tuple[Optional[str], Optional[str]]:
    """Show a GUI popup to select input Excel file and output .asy file."""

    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
    )

    if not excel_file:
        print("No input file selected. Exiting.")
        root.destroy()
        return None, None

    output_file = filedialog.asksaveasfilename(
        title="Save .asy File As",
        defaultextension=".asy",
        filetypes=[
            ("ASY files", "*.asy"),
            ("All files", "*.*"),
        ],
        initialfile=Path(excel_file).stem + ".asy",
    )

    if not output_file:
        print("No output file selected. Exiting.")
        root.destroy()
        return None, None

    root.destroy()
    return excel_file, output_file


def show_file_dialog_notebook() -> Tuple[Optional[str], Optional[str]]:
    """Show a GUI popup to select input Excel file and output notebook file."""

    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
    )

    if not excel_file:
        print("No input file selected. Exiting.")
        root.destroy()
        return None, None

    output_file = filedialog.asksaveasfilename(
        title="Save Notebook File As",
        defaultextension=".ipynb",
        filetypes=[
            ("Jupyter Notebook", "*.ipynb"),
            ("All files", "*.*"),
        ],
        initialfile=Path(excel_file).stem + "_liquid_handler.ipynb",
    )

    if not output_file:
        print("No output file selected. Exiting.")
        root.destroy()
        return None, None

    root.destroy()
    return excel_file, output_file


def main() -> None:
    """Entry point for command-line and GUI usage."""

    if len(sys.argv) > 1:
        parser = argparse.ArgumentParser(
            description=(
                "Convert GatorBio Assay Form Excel file to .asy file or generate liquid handler notebook"
            )
        )
        parser.add_argument(
            "excel_file",
            help="Path to the Excel file (.xlsx or .xlsm)",
        )
        parser.add_argument(
            "-o",
            "--output",
            help="Output file path (default: same name as Excel file with appropriate extension)",
        )
        parser.add_argument(
            "--notebook",
            "-n",
            action="store_true",
            help="Generate liquid handler notebook instead of .asy file",
        )

        args = parser.parse_args()

        try:
            if args.notebook:
                generate_liquid_handler_notebook(args.excel_file, args.output)
            else:
                generate_asy_file(args.excel_file, args.output)
        except Exception as exc:  # noqa: BLE001
            print(f"Error: {exc}", file=sys.stderr)
            import traceback

            traceback.print_exc()
            sys.exit(1)
        return

    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        choice = messagebox.askyesnocancel(
            "Select Output Type",
            (
                "What would you like to generate?\n\n"
                "Yes = .asy file (GatorBio assay file)\n"
                "No = .ipynb file (Liquid handler notebook)\n"
                "Cancel = Exit"
            ),
        )

        if choice is None:
            root.destroy()
            return
        if choice:
            excel_file, output_file = show_file_dialog()
            if excel_file and output_file:
                generate_asy_file(excel_file, output_file)
                messagebox.showinfo("Success", f"Successfully generated:\n{output_file}")
        else:
            excel_file, output_file = show_file_dialog_notebook()
            if excel_file and output_file:
                generate_liquid_handler_notebook(excel_file, output_file)
                messagebox.showinfo("Success", f"Successfully generated:\n{output_file}")

        root.destroy()
    except Exception as exc:  # noqa: BLE001
        import traceback

        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror("Error", f"An error occurred:\n{str(exc)}")
        root.destroy()

        print(f"Error: {exc}", file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)


__all__ = ["main", "show_file_dialog", "show_file_dialog_notebook"]

