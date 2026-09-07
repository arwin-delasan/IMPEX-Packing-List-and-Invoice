#!/usr/bin/env python3
"""GUI for generating a categorized SAWO packing list and/or invoice.

Pick a SAWO packing-list PDF once - this extracts it, looks up each item in
Odoo, and classifies it by product code/name into the Sauna Equipment /
Accessories / Spare Parts / etc. groups (categorize.py). From there, use
"Generate Packing List..." and/or "Generate Invoice..." as needed - both
render the same extracted/classified data, just with different columns and
totals (see generate_pl1.py / generate_inv.py), so you're not re-picking or
re-classifying the PDF for each output.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from pdf_parser import ExtractionError
from pdf_parser import extract_pdf as extract_pdf_custom
from odoo_pdf_parser import extract_pdf as extract_pdf_odoo
from generate_pl1 import (
    GenerationError,
    build_blocks,
    build_workbook,
    classify_items,
    um_warnings,
    validate,
)
from generate_inv import apply_bella_vivo_billing, build_invoice_workbook, load_bella_vivo_prices
from proforma_parser import parse_proforma
from odoo_client import OdooClient, OdooError, _load_env
from categorize import UNCATEGORIZED, CATEGORY_ORDER
from overrides import OverridesError, delete_override, load_overrides, set_overrides_bulk
from app_paths import app_path


def login_dialog(parent, odoo_url, odoo_db):
    """Modal Odoo login prompt shown once at startup. odoo_url/odoo_db are
    shared (same server/database for everyone, read from .env) and shown
    read-only for context; username/password are entered fresh each
    session instead of living in .env, so one shared install (e.g. the
    LAN-distributed .exe) doesn't bake one person's login into every
    copy. Returns a connected OdooClient, or None if the window was
    closed without logging in successfully."""
    result = {"client": None}

    dialog = tk.Toplevel(parent)
    dialog.title("Odoo Login")
    dialog.resizable(False, False)
    dialog.transient(parent)

    tk.Label(
        dialog, text=f"Server: {odoo_url}\nDatabase: {odoo_db}",
        justify="left", fg="#555555",
    ).pack(anchor="w", padx=20, pady=(16, 8))

    form = tk.Frame(dialog)
    form.pack(padx=20, pady=(0, 4))
    tk.Label(form, text="Username:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=4)
    username_var = tk.StringVar()
    username_entry = tk.Entry(form, textvariable=username_var, width=28)
    username_entry.grid(row=0, column=1, pady=4)
    tk.Label(form, text="Password:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=4)
    password_var = tk.StringVar()
    password_entry = tk.Entry(form, textvariable=password_var, width=28, show="*")
    password_entry.grid(row=1, column=1, pady=4)

    error_label = tk.Label(dialog, text="", fg="#b00020", wraplength=280, justify="left")
    error_label.pack(fill="x", padx=20)

    def attempt_login(event=None):
        username = username_var.get().strip()
        password = password_var.get()
        if not username or not password:
            error_label.configure(text="Enter both username and password.")
            return
        login_btn.configure(state="disabled", text="Connecting...")
        error_label.configure(text="")
        dialog.update_idletasks()
        try:
            client = OdooClient(username=username, password=password)
        except OdooError as e:
            error_label.configure(text=str(e))
            login_btn.configure(state="normal", text="Log In")
            password_entry.focus_set()
            return
        result["client"] = client
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=(8, 16))
    login_btn = tk.Button(btn_frame, text="Log In", width=14, command=attempt_login)
    login_btn.pack()

    username_entry.focus_set()
    dialog.bind("<Return>", attempt_login)

    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    dialog.update_idletasks()
    x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
    y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")

    dialog.grab_set()
    parent.wait_window(dialog)
    return result["client"]


def ask_choice(parent, title, message, options):
    """Modal dialog with one labeled button per option (no generic Yes/No/
    Cancel - each button reads exactly like its option, e.g. "Odoo PDF" /
    "Email PDF"). Returns the chosen option string, or None if the dialog
    was closed without choosing (e.g. via the window's close button)."""
    result = {"value": None}

    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.resizable(False, False)
    dialog.transient(parent)

    tk.Label(dialog, text=message, justify="left", padx=20, pady=16, wraplength=360).pack()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(padx=20, pady=(0, 16))

    def choose(value):
        result["value"] = value
        dialog.destroy()

    for option in options:
        tk.Button(btn_frame, text=option, width=14, command=lambda v=option: choose(v)).pack(side="left", padx=6)

    dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
    dialog.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - dialog.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - dialog.winfo_height()) // 2
    dialog.geometry(f"+{x}+{y}")

    dialog.grab_set()
    parent.wait_window(dialog)
    return result["value"]


class App:
    def __init__(self, root):
        self.root = root
        # Most recent PDF's extracted/classified state, kept in memory only.
        self._last_pdf_path = None
        self._last_header = None
        self._last_items = None
        self._last_grand_total = None
        self._last_classified = None

        root.title("SAWO Packing List / Invoice Generator")
        root.geometry("600x520")
        root.minsize(480, 400)
        # Deliberately not root.withdraw() here - a Toplevel that's
        # .transient() to a withdrawn master inherits that withdrawn
        # state itself in this Tk build (confirmed: its .state() comes
        # back "withdrawn" even after an explicit .deiconify()), so the
        # login dialog below would silently never actually show. Root
        # just appears briefly blank behind the (modal, grab_set) login
        # dialog instead - harmless.

        try:
            env = _load_env(app_path(".env"))
        except OSError:
            env = {}
        self.odoo = login_dialog(root, env.get("ODOO_URL", "(not set in .env)"), env.get("ODOO_DB", "(not set in .env)"))
        if self.odoo is None:
            return  # main() checks self.odoo and closes the app without starting mainloop

        frame = tk.Frame(root, padx=12, pady=12)
        frame.pack(fill="both", expand=True)

        self.choose_btn = tk.Button(
            frame, text="Choose Packing List PDF...", height=2, command=self.run_flow
        )
        self.choose_btn.pack(fill="x", pady=(0, 10))

        output_frame = tk.Frame(frame)
        output_frame.pack(fill="x", pady=(0, 10))
        self.packing_list_btn = tk.Button(
            output_frame, text="Generate Packing List...", height=2,
            command=self.generate_packing_list, state="disabled",
        )
        self.packing_list_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.invoice_btn = tk.Button(
            output_frame, text="Generate Invoice...", height=2,
            command=self.generate_invoice, state="disabled",
        )
        self.invoice_btn.pack(side="left", fill="x", expand=True, padx=(5, 0))

        self.review_btn = tk.Button(
            frame, text="Review & Correct Classification...", height=2,
            command=self.open_review_window, state="disabled",
        )
        self.review_btn.pack(fill="x", pady=(0, 10))

        self.manage_btn = tk.Button(
            frame, text="Manage Corrections...", height=2, command=self.open_manage_window
        )
        self.manage_btn.pack(fill="x", pady=(0, 10))

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

        # Two real, differently-laid-out packing-list templates exist -
        # the original custom one (emailed as an attachment), and one
        # rendered by Odoo's own report engine (different header layout,
        # different item-table column order) - so ask which extractor to
        # use rather than guessing.
        choice = ask_choice(
            self.root, "PDF format", "Which kind of Packing List PDF is this?",
            ["Odoo PDF", "Email PDF"],
        )
        if choice is None:
            return
        extract_pdf = extract_pdf_odoo if choice == "Odoo PDF" else extract_pdf_custom

        self.clear_log()
        self.choose_btn.configure(state="disabled")
        self.packing_list_btn.configure(state="disabled")
        self.invoice_btn.configure(state="disabled")
        threading.Thread(target=self._run_flow, args=(pdf_path, extract_pdf), daemon=True).start()

    def _run_flow(self, pdf_path, extract_pdf):
        try:
            self.root.after(0, self.write_log, f"Reading PDF: {pdf_path}")
            header, items, grand_total = extract_pdf(pdf_path)
            self.root.after(0, self.write_log, f"Extracted header block and {len(items)} line items.")
            for w in um_warnings(items):
                self.root.after(0, self.write_log, f"WARNING: {w}")

            classified, warnings = classify_items(items, self.odoo)
            for w in warnings:
                self.root.after(0, self.write_log, f"WARNING: {w}")
            self.root.after(0, self.write_log, f"Classified {len(classified)} items ({len(warnings)} warning(s)).")

            totals_warnings = validate(items, grand_total)
            if not totals_warnings:
                self.root.after(0, self.write_log, "Validation passed: extracted totals reconcile with PDF Grand Total.")
            else:
                for w in totals_warnings:
                    self.root.after(0, self.write_log, f"WARNING: {w}")

            self._last_pdf_path = pdf_path
            self._last_header = header
            self._last_items = items
            self._last_grand_total = grand_total
            self._last_classified = classified

            self.root.after(0, lambda: self.packing_list_btn.configure(state="normal"))
            self.root.after(0, lambda: self.invoice_btn.configure(state="normal"))
            self.root.after(0, lambda: self.review_btn.configure(state="normal"))
            self.root.after(
                0, self.write_log,
                'Ready — use "Generate Packing List..." and/or "Generate Invoice..." below.',
            )
        except (ExtractionError, GenerationError, OdooError, OverridesError) as e:
            self.root.after(0, self.write_log, f"ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Extraction failed", str(e))
        except Exception as e:
            self.root.after(0, self.write_log, f"UNEXPECTED ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Unexpected error", str(e))
        finally:
            self.root.after(0, lambda: self.choose_btn.configure(state="normal"))

    def generate_packing_list(self):
        self._generate_output(
            label="Packing List", suffix="Packing List",
            build_fn=build_workbook, output_desc="categorized packing list",
        )

    def generate_invoice(self):
        # Which price source depends on the customer, not anything we can
        # detect automatically - just ask.
        choice = ask_choice(
            self.root, "Pricing source", "Which pricing source should this invoice use?",
            ["Bella Vivo", "Proforma"],
        )
        if choice is None:
            return
        currency = "USD"
        if choice == "Bella Vivo":
            price_lookup = load_bella_vivo_prices()
            header = apply_bella_vivo_billing(self._last_header)
        else:
            proforma_path = filedialog.askopenfilename(
                title="Choose this order's Pro-Forma Invoice PDF",
                initialdir=os.path.dirname(self._last_pdf_path),
                filetypes=[("PDF files", "*.pdf")],
            )
            if not proforma_path:
                return
            try:
                price_lookup, currency = parse_proforma(proforma_path)
            except ExtractionError as e:
                messagebox.showerror("Pro-Forma Invoice extraction failed", str(e))
                return
            self.write_log(f"Read {len(price_lookup)} unit price(s) from the Pro-Forma Invoice ({currency}).")
            header = self._last_header

        self._generate_output(
            label="Invoice", suffix="Invoice",
            build_fn=build_invoice_workbook, output_desc="invoice",
            header=header, extra_kwargs={"price_lookup": price_lookup, "currency": currency},
        )

    def _generate_output(self, label, suffix, build_fn, output_desc, header=None, extra_kwargs=None):
        if not self._last_classified:
            return

        default_name = f"{os.path.splitext(os.path.basename(self._last_pdf_path))[0]} {suffix}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title=f"Save {output_desc} as...",
            initialdir=os.path.dirname(self._last_pdf_path),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not output_path:
            return

        self.packing_list_btn.configure(state="disabled")
        self.invoice_btn.configure(state="disabled")
        threading.Thread(
            target=self._generate_output_bg,
            args=(label, build_fn, output_path, header or self._last_header, extra_kwargs or {}), daemon=True,
        ).start()

    def _generate_output_bg(self, label, build_fn, output_path, header, extra_kwargs):
        try:
            # Rebuilt fresh from self._last_classified each time (not
            # cached), so any correction made via "Review & Correct
            # Classification..." since the PDF was picked is reflected.
            blocks = build_blocks(self._last_classified)
            self.root.after(0, self.write_log, f"Grouped into {len(blocks)} category/subtype block(s).")

            result = build_fn(header, blocks, output_path, **extra_kwargs)
            for w in result or []:  # e.g. codes missing from a price list
                self.root.after(0, self.write_log, f"WARNING: {w}")
            self.root.after(0, self.write_log, f"Wrote {label.lower()}: {output_path}")
            self.root.after(0, messagebox.showinfo, "Done", f"{label} generated.\n\nSaved to:\n{output_path}")
        except (ExtractionError, GenerationError, OdooError, OverridesError) as e:
            self.root.after(0, self.write_log, f"ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Generation failed", str(e))
        except Exception as e:
            self.root.after(0, self.write_log, f"UNEXPECTED ERROR: {e}")
            self.root.after(0, messagebox.showerror, "Unexpected error", str(e))
        finally:
            self.root.after(0, lambda: self.packing_list_btn.configure(state="normal"))
            self.root.after(0, lambda: self.invoice_btn.configure(state="normal"))

    def open_review_window(self):
        if not self._last_classified:
            return

        win = tk.Toplevel(self.root)
        win.title("Review & Correct Classification")
        win.geometry("900x500")

        tk.Label(
            win,
            text="Corrections apply to future runs only — this file has already been "
                 "saved and won't change.",
            anchor="w", justify="left", wraplength=860,
        ).pack(fill="x", padx=10, pady=(10, 6))

        search_frame = tk.Frame(win)
        search_frame.pack(fill="x", padx=10, pady=(0, 6))
        tk.Label(search_frame, text="Search Code:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_frame, textvariable=search_var, width=30)
        search_entry.pack(side="left", padx=(6, 6))
        tk.Button(search_frame, text="Clear", command=lambda: search_var.set("")).pack(side="left")

        columns = ("code", "name", "category", "subtype", "status")
        headings = {
            "code": "Code", "name": "Product Name",
            "category": "Category", "subtype": "Subtype", "status": "Status",
        }
        widths = {"code": 110, "name": 340, "category": 150, "subtype": 150, "status": 90}

        tree = ttk.Treeview(win, columns=columns, show="headings", selectmode="extended")
        for col in columns:
            tree.column(col, width=widths[col], anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        tree.tag_configure("uncategorized", background="#fddede")

        # Click a column header to sort by it; click again to reverse.
        sort_state = {"col": None, "reverse": False}
        sort_keys = {
            "code": lambda it: it["external_code"],
            "name": lambda it: it["name"],
            "category": lambda it: it["category"],
            "subtype": lambda it: it["subtype"] or "",
            "status": lambda it: "Corrected" if it.get("overridden") else "",
        }

        def sort_by(col):
            if sort_state["col"] == col:
                sort_state["reverse"] = not sort_state["reverse"]
            else:
                sort_state["col"] = col
                sort_state["reverse"] = False
            populate()

        def refresh_headings():
            for col in columns:
                text = headings[col]
                if sort_state["col"] == col:
                    text += " ▼" if sort_state["reverse"] else " ▲"
                tree.heading(col, text=text, command=lambda c=col: sort_by(c))

        def populate():
            refresh_headings()
            tree.delete(*tree.get_children())
            entries = list(enumerate(self._last_classified))  # (original_index, item)
            query = search_var.get().strip().lower()
            if query:
                entries = [e for e in entries if query in e[1]["external_code"].lower()]
            if sort_state["col"]:
                entries.sort(key=lambda e: sort_keys[sort_state["col"]](e[1]), reverse=sort_state["reverse"])
            for i, it in entries:
                tags = ("uncategorized",) if it["category"] == UNCATEGORIZED else ()
                tree.insert("", "end", iid=str(i), tags=tags, values=(
                    it["external_code"], it["name"],
                    it["category"], it["subtype"] or "",
                    "Corrected" if it.get("overridden") else "",
                ))

        search_var.trace_add("write", lambda *args: populate())
        populate()

        def correct_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No selection", "Select one or more rows to correct.", parent=win)
                return
            self.open_correction_dialog(win, [int(iid) for iid in selected], populate)

        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btn_frame, text="Correct Selected...", command=correct_selected).pack(side="left")
        tk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right")

    def open_correction_dialog(self, parent, indices, on_saved):
        codes = []
        seen = set()
        for i in indices:
            code = self._last_classified[i]["external_code"]
            if code not in seen:
                seen.add(code)
                codes.append(code)

        summary = f"Correcting {len(indices)} item(s), {len(codes)} unique code(s): " + ", ".join(codes[:8])
        if len(codes) > 8:
            summary += f", +{len(codes) - 8} more"
        known_subtypes = sorted({it["subtype"] for it in self._last_classified if it["subtype"]})

        def on_confirm(category, subtype):
            product_names = {}
            originals = {}
            for i in indices:
                it = self._last_classified[i]
                code = it["external_code"]
                product_names[code] = it["name"]
                originals[code] = (it["category"], it["subtype"])

            try:
                set_overrides_bulk(codes, category, subtype, product_names, originals)
            except (ValueError, OverridesError) as e:
                messagebox.showerror("Could not save correction", str(e), parent=parent)
                return False

            for i in indices:
                self._last_classified[i]["category"] = category
                self._last_classified[i]["subtype"] = subtype
                self._last_classified[i]["overridden"] = True

            on_saved()
            return True

        self._open_category_subtype_picker(
            parent, "Correct Classification", summary, known_subtypes,
            initial_category=None, initial_subtype=None, on_confirm=on_confirm,
        )

    def open_edit_override_dialog(self, parent, codes, current_overrides, on_saved):
        """Edit one or more already-saved corrections directly, from the
        Manage Corrections window - no prior run/self._last_classified
        needed, since the data comes from the saved overrides themselves."""
        first = current_overrides[codes[0]]
        summary = f"Editing {len(codes)} correction(s): " + ", ".join(codes[:8])
        if len(codes) > 8:
            summary += f", +{len(codes) - 8} more"
        known_subtypes = sorted({
            e.get("subtype") for e in current_overrides.values() if e.get("subtype")
        })

        def on_confirm(category, subtype):
            product_names = {code: current_overrides[code].get("product_name", "") for code in codes}
            originals = {
                code: (current_overrides[code].get("category"), current_overrides[code].get("subtype"))
                for code in codes
            }
            try:
                set_overrides_bulk(codes, category, subtype, product_names, originals)
            except (ValueError, OverridesError) as e:
                messagebox.showerror("Could not save correction", str(e), parent=parent)
                return False
            on_saved()
            return True

        self._open_category_subtype_picker(
            parent, "Edit Correction", summary, known_subtypes,
            initial_category=first.get("category"), initial_subtype=first.get("subtype"),
            on_confirm=on_confirm,
        )

    def _open_category_subtype_picker(
        self, parent, title, summary_text, known_subtypes,
        initial_category, initial_subtype, on_confirm,
    ):
        """Shared Category/Subtype entry dialog used by both "Correct
        Selected" (Review window) and "Edit Selected" (Manage Corrections
        window). on_confirm(category, subtype) does the actual saving and
        returns True to close the dialog, or False to leave it open (after
        showing its own error) so the user can fix and retry."""
        dialog = tk.Toplevel(parent)
        dialog.title(title)
        dialog.transient(parent)
        dialog.grab_set()

        tk.Label(dialog, text=summary_text, wraplength=380, justify="left").grid(
            row=0, column=0, columnspan=2, padx=10, pady=(10, 6), sticky="w"
        )

        tk.Label(dialog, text="Category:").grid(row=1, column=0, padx=10, pady=4, sticky="w")
        category_var = tk.StringVar(value=initial_category or "")
        category_box = ttk.Combobox(dialog, textvariable=category_var, values=CATEGORY_ORDER, width=30)
        category_box.grid(row=1, column=1, padx=10, pady=4, sticky="w")

        def filter_categories(event=None):
            typed = category_var.get()
            if not typed:
                category_box["values"] = CATEGORY_ORDER
                return
            filtered = [c for c in CATEGORY_ORDER if typed.lower() in c.lower()]
            category_box["values"] = filtered or CATEGORY_ORDER

        category_box.bind("<KeyRelease>", filter_categories)

        tk.Label(
            dialog, text="Not in the list? Just type a new one.",
            font=("TkDefaultFont", 8), fg="gray",
        ).grid(row=1, column=2, padx=(0, 10), pady=4, sticky="w")

        tk.Label(dialog, text="Subtype:").grid(row=2, column=0, padx=10, pady=4, sticky="w")
        subtype_var = tk.StringVar(value=initial_subtype or "(none)")
        subtype_values = ["(none)"] + known_subtypes
        subtype_box = ttk.Combobox(dialog, textvariable=subtype_var, values=subtype_values, width=30)
        subtype_box.grid(row=2, column=1, padx=10, pady=4, sticky="w")

        def filter_subtypes(event=None):
            typed = subtype_var.get()
            if not typed or typed == "(none)":
                subtype_box["values"] = subtype_values
                return
            filtered = [s for s in subtype_values if typed.lower() in s.lower()]
            subtype_box["values"] = filtered or subtype_values

        subtype_box.bind("<KeyRelease>", filter_subtypes)

        def on_ok():
            category = category_var.get().strip()
            if not category:
                messagebox.showwarning("Missing category", "Choose or type a category.", parent=dialog)
                return
            if category not in CATEGORY_ORDER:
                if not messagebox.askyesno(
                    "New category",
                    f'"{category}" isn\'t one of the existing categories:\n\n'
                    + "\n".join(CATEGORY_ORDER)
                    + '\n\nCreate it as a new category? (Check for typos first — '
                      'if this was meant to match an existing one, click No.)',
                    parent=dialog,
                ):
                    return
            subtype_raw = subtype_var.get().strip()
            subtype = None if subtype_raw in ("", "(none)") else subtype_raw

            if on_confirm(category, subtype):
                dialog.destroy()

        btns = tk.Frame(dialog)
        btns.grid(row=3, column=0, columnspan=2, pady=(10, 10))
        tk.Button(btns, text="OK", width=10, command=on_ok).pack(side="left", padx=6)
        tk.Button(btns, text="Cancel", width=10, command=dialog.destroy).pack(side="left", padx=6)

    def open_manage_window(self):
        win = tk.Toplevel(self.root)
        win.title("Manage Corrections")
        win.geometry("1000x450")

        columns = ("code", "name", "category_change", "subtype_change", "corrected_at")
        headings = {
            "code": "Code", "name": "Product Name", "category_change": "Category Change",
            "subtype_change": "Subtype Change", "corrected_at": "Corrected At",
        }
        widths = {"code": 110, "name": 280, "category_change": 220, "subtype_change": 220, "corrected_at": 140}

        tree = ttk.Treeview(win, columns=columns, show="headings", selectmode="extended")
        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def fmt(old, new):
            return f"{old if old else '(none)'} → {new if new else '(none)'}"

        def populate():
            tree.delete(*tree.get_children())
            try:
                data = load_overrides()
            except OverridesError as e:
                messagebox.showerror("Could not load corrections", str(e), parent=win)
                return
            for code, entry in sorted(data.items(), key=lambda kv: kv[1].get("corrected_at", ""), reverse=True):
                tree.insert("", "end", iid=code, values=(
                    code, entry.get("product_name", ""),
                    fmt(entry.get("original_category"), entry.get("category")),
                    fmt(entry.get("original_subtype"), entry.get("subtype")),
                    entry.get("corrected_at", ""),
                ))

        populate()

        def edit_selected():
            selected = tree.selection()
            if not selected:
                messagebox.showinfo("No selection", "Select one or more corrections to edit.", parent=win)
                return
            try:
                current = load_overrides()
            except OverridesError as e:
                messagebox.showerror("Could not load corrections", str(e), parent=win)
                return
            self.open_edit_override_dialog(win, list(selected), current, populate)

        def delete_selected():
            selected = tree.selection()
            if not selected:
                return
            if not messagebox.askyesno(
                "Delete corrections",
                f"Delete {len(selected)} saved correction(s)? This can't be undone.",
                parent=win,
            ):
                return
            for code in selected:
                delete_override(code)
            populate()

        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btn_frame, text="Refresh", command=populate).pack(side="left")
        tk.Button(btn_frame, text="Edit Selected...", command=edit_selected).pack(side="left", padx=(6, 0))
        tk.Button(btn_frame, text="Delete Selected", command=delete_selected).pack(side="left", padx=(6, 0))
        tk.Button(btn_frame, text="Close", command=win.destroy).pack(side="right")


def main():
    root = tk.Tk()
    app = App(root)
    if app.odoo is None:  # login was cancelled - nothing to run without it
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
