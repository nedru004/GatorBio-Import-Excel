#!/usr/bin/env python3
"""
Entry point for the GatorBio Excel conversion utilities.

Provides command-line and GUI access for generating .asy files and liquid handler
notebooks from the GatorBio Excel assay form.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional, Tuple

from gatorbio_excel.asy_generator import generate_asy_file
from gatorbio_excel.notebook_generator import generate_liquid_handler_notebook


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


def show_file_dialog_notebook() -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str]]:
    """Show a GUI popup to select input Excel file, output notebook file, buffer source, and no-dilution handling."""

    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.title("Liquid Handler Notebook Options")
    root.attributes("-topmost", True)
    root.geometry("400x250")

    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
        parent=root,
    )

    if not excel_file:
        print("No input file selected. Exiting.")
        root.destroy()
        return None, None, None, None

    output_file = filedialog.asksaveasfilename(
        title="Save Notebook File As",
        defaultextension=".ipynb",
        filetypes=[
            ("Jupyter Notebook", "*.ipynb"),
            ("All files", "*.*"),
        ],
        initialfile=Path(excel_file).stem + "_liquid_handler.ipynb",
        parent=root,
    )

    if not output_file:
        print("No output file selected. Exiting.")
        root.destroy()
        return None, None, None, None

    # Create options dialog
    dialog = tk.Toplevel(root)
    dialog.title("Liquid Handler Notebook Options")
    dialog.attributes("-topmost", True)
    dialog.geometry("450x200")
    dialog.transient(root)
    dialog.grab_set()

    options_frame = tk.Frame(dialog, padx=20, pady=20)
    options_frame.pack(fill=tk.BOTH, expand=True)

    # Buffer source selection
    tk.Label(options_frame, text="Buffer Source:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=10)
    buffer_var = tk.StringVar(value="trough")
    buffer_frame = tk.Frame(options_frame)
    buffer_frame.grid(row=0, column=1, sticky="w", pady=10)
    tk.Radiobutton(buffer_frame, text="50 mL Tube", variable=buffer_var, value="tube").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(buffer_frame, text="60 mL Trough", variable=buffer_var, value="trough").pack(side=tk.LEFT, padx=5)

    # No-dilution samples handling
    tk.Label(options_frame, text="Samples with no dilution:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=10)
    no_dilution_var = tk.StringVar(value="pipette")
    no_dilution_frame = tk.Frame(options_frame)
    no_dilution_frame.grid(row=1, column=1, sticky="w", pady=10)
    tk.Radiobutton(no_dilution_frame, text="Pipette automatically", variable=no_dilution_var, value="pipette").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(no_dilution_frame, text="Skip (manual addition)", variable=no_dilution_var, value="skip").pack(side=tk.LEFT, padx=5)

    # OK/Cancel buttons
    button_frame = tk.Frame(options_frame)
    button_frame.grid(row=2, column=0, columnspan=2, pady=20)
    
    result = {"confirmed": False, "buffer_source": None, "no_dilution_handling": None}
    
    def on_ok():
        result["confirmed"] = True
        result["buffer_source"] = buffer_var.get()
        result["no_dilution_handling"] = no_dilution_var.get()
        dialog.destroy()
    
    def on_cancel():
        result["confirmed"] = False
        dialog.destroy()
    
    tk.Button(button_frame, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Cancel", command=on_cancel, width=10).pack(side=tk.LEFT, padx=5)

    dialog.wait_window()
    root.destroy()
    
    if not result["confirmed"]:
        return None, None, None, None

    return excel_file, output_file, result["buffer_source"], result["no_dilution_handling"]


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
        parser.add_argument(
            "--buffer-source",
            choices=["tube", "trough"],
            default="tube",
            help="Buffer source type: 'tube' for 50mL tube or 'trough' for 60mL trough (default: tube)",
        )
        parser.add_argument(
            "--no-dilution-handling",
            choices=["pipette", "skip"],
            default="pipette",
            help="How to handle samples with no dilution: 'pipette' to pipette automatically, 'skip' to skip (manual addition) (default: pipette)",
        )

        args = parser.parse_args()

        try:
            if args.notebook:
                generate_liquid_handler_notebook(
                    args.excel_file, 
                    args.output, 
                    buffer_source=args.buffer_source,
                    skip_no_dilution_samples=(args.no_dilution_handling == "skip")
                )
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
            result = show_file_dialog_notebook()
            if result and result[0] and result[1] and result[2] and result[3]:
                excel_file, output_file, buffer_source, no_dilution_handling = result
                generate_liquid_handler_notebook(
                    excel_file, 
                    output_file, 
                    buffer_source=buffer_source,
                    skip_no_dilution_samples=(no_dilution_handling == "skip")
                )
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


if __name__ == "__main__":
    main()
