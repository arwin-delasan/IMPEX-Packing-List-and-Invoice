#!/usr/bin/env python3
"""GUI for generating a categorized SAWO packing list.

Pick a SAWO packing-list PDF, and this extracts it, looks up each item in
Odoo, classifies it by product code/name into the Sauna Equipment /
Accessories / Spare Parts / etc. groups used by the historical "PL1"
packing-list format, and writes a grouped, subtotaled workbook styled
after "PL1 - Copy.xlsx" (expected next to this script).

Pick the PDF, then a Save dialog lets you choose where the output goes
(defaults to the PDF's own name with .xlsx).
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from pdf_parser import ExtractionError, extract_pdf
from generate_pl1 import (
    GenerationError,
    build_blocks,
    build_workbook,
    classify_items,
    validate,
)
from odoo_client import OdooClient, OdooError


class App:
    def __init__(self, root):
        self.root = root

        root.title("SAWO Packing List Generator")
        root.geometry("600x420")
        root.minsize(480, 320)

        frame = tk.Frame(root, padx=12, pady=12)
        frame.pack(fill="both", expand=True)

        self.choose_btn = tk.Button(
            frame, text="Choose Packing List PDF...", height=2, command=self.run_flow
        )
        self.choose_btn.pack(fill="x", pady=(0, 10))

        self.log = scrolledtext.ScrolledText(frame, height=16, state="disabled")
        self.log.pack(fill="both", expand=True)

    def write_log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def run_flow(self):
        pdf_path = filedialog.askopenfilename(
            title="Choose packing list PDF",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            return

        default_name = os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Save categorized packing list as...",
            initialdir=os.path.dirname(pdf_path),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output_path:
            return

        self.clear_log()
        self.choose_btn.configure(state="disabled")
        threading.Thread(
            target=self._run_flow, args=(pdf_path, output_path), daemon=True
        ).start()

    def _run_flow(self, pdf_path, output_path):
        try:
            self.root.after(0, self.write_log, f"Reading PDF: {pdf_path}")
            header, items, grand_total = extract_pdf(pdf_path)
            self.root.after(0, self.write_log, f"Extracted header block and {len(items)} line items.")

            self.root.after(0, self.write_log, "Connecting to Odoo...")
            odoo = OdooClient()

            classified, warnings = classify_items(items, odoo)
            for w in warnings:
                self.root.after(0, self.write_log, f"WARNING: {w}")
            self.root.after(0, self.write_log, f"Classified {len(classified)} items ({len(warnings)} warning(s)).")

            blocks = build_blocks(classified)
            self.root.after(0, self.write_log, f"Grouped into {len(blocks)} category/subtype block(s).")

            build_workbook(header, blocks, output_path, odoo)
            self.root.after(0, self.write_log, f"Wrote output workbook: {output_path}")

            totals_warnings = validate(items, grand_total)
            if not totals_warnings:
                self.root.after(0, self.write_log, "Validation passed: totals reconcile with PDF Grand Total.")
                self.root.after(0, messagebox.showinfo, "Done", f"Generation complete.\n\nSaved to:\n{output_path}")
            else:
                for w in totals_warnings:
                    self.root.after(0, self.write_log, f"WARNING: {w}")
                self.root.after(
                    0, messagebox.showwarning, "Validation warning",
                    "File was saved, but totals did not reconcile with the PDF "
                    f"Grand Total. See the log for details.\n\nSaved to:\n{output_path}",
                )
        except (ExtractionError, GenerationError, OdooError) as e:
            self.root.after(0, self.write_log, f"ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Generation failed", str(e))
        except Exception as e:
            self.root.after(0, self.write_log, f"UNEXPECTED ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Unexpected error", str(e))
        finally:
            self.root.after(0, lambda: self.choose_btn.configure(state="normal"))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
