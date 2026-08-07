"""
Sage File Organiser - Desktop Application
Processes and reorganizes Sage CSV exports (Customers, Sales Orders/Invoices, Products, Transactions).
Re-orders columns, renames headers, cleans phone/address data, normalizes dates, and performs data segmentation.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import os
import re
from datetime import datetime, timedelta
import threading


# ═══════════════════════════════════════════════════════════════
# Encoding Helpers & File I/O
# ═══════════════════════════════════════════════════════════════

ENCODINGS_TO_TRY = ["utf-8-sig", "cp1252", "latin-1", "iso-8859-1"]


def open_csv_file(file_path):
    """Open a CSV file trying multiple encodings.

    Returns:
        Tuple of (open_file_handle, encoding_used)
    """
    for enc in ENCODINGS_TO_TRY:
        try:
            f = open(file_path, "r", newline="", encoding=enc)
            f.read()
            f.seek(0)
            return f, enc
        except UnicodeDecodeError:
            try:
                f.close()
            except Exception:
                pass
            continue
        except Exception:
            try:
                f.close()
            except Exception:
                pass
            raise

    error_context = _get_decode_error_context(file_path)
    raise ValueError(
        f"Could not decode file with any supported encoding ({', '.join(ENCODINGS_TO_TRY)}).\n\n{error_context}"
    )


def _get_decode_error_context(file_path):
    """Read raw bytes and generate diagnostic context for decoding errors."""
    try:
        with open(file_path, "rb") as f:
            raw = f.read()
        try:
            raw.decode("utf-8")
        except UnicodeDecodeError as e:
            pos = e.start
            start = max(0, pos - 100)
            before_text = raw[start:pos].decode("latin-1", errors="replace")
            after_text = raw[pos:pos + 20].decode("latin-1", errors="replace")
            bad_byte = raw[pos:pos + 1]
            return (
                f"Decode error at byte position {pos}.\n"
                f"Problem byte: {bad_byte!r} (hex: {bad_byte.hex()})\n\n"
                f"── Text before error ──\n{before_text}\n\n"
                f"── Text after error ──\n{after_text}"
            )
        return "File could not be decoded."
    except Exception as e:
        return f"Could not read raw file bytes: {e}"


def clean_phone_number(phone_str):
    """Format phone numbers cleanly, handling trailing .0 and leading zero formatting."""
    if not phone_str:
        return ""
    val = str(phone_str).strip()
    if val.endswith(".0"):
        val = val[:-2]
    # Remove unwanted punctuation except + and digits
    val = re.sub(r"[^\d+]", "", val)
    if val and not val.startswith("+") and not val.startswith("0") and len(val) >= 10:
        val = "0" + val
    return val


def format_date(date_str):
    """Attempt to parse and normalize date string to YYYY-MM-DD format."""
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    patterns = [
        "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
        "%d/%m/%y", "%d-%m-%y", "%Y/%m/%d"
    ]
    for fmt in patterns:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return date_str


# ═══════════════════════════════════════════════════════════════
# Sage Default Mappings & Schemas
# ═══════════════════════════════════════════════════════════════

SAGE_CUSTOMER_MAP = {
    "account_ref": "Customer_ID",
    "name": "Customer_Name",
    "address_1": "Address_Line_1",
    "address_2": "Address_Line_2",
    "address_3": "Address_Line_3",
    "address_4": "Address_Line_4",
    "address_5": "Address_Line_5",
    "postcode": "Postcode",
    "telephone": "Phone",
    "telephone_2": "Phone_2",
    "contact_name": "Contact_Name",
    "e_mail": "Email",
    "balance": "Current_Balance",
    "credit_limit": "Credit_Limit",
}

SAGE_CUSTOMER_OUTPUT_ORDER = [
    "Customer_ID", "Customer_Name", "Contact_Name", "Email", "Phone", "Phone_2",
    "Address_Line_1", "Address_Line_2", "Address_Line_3", "Address_Line_4", "Address_Line_5",
    "Postcode", "Current_Balance", "Credit_Limit", "Segment", "Status"
]

SAGE_ORDER_MAP = {
    "order_number": "Order_Number",
    "account_ref": "Customer_ID",
    "order_date": "Order_Date",
    "cust_order_number": "PO_Reference",
    "stock_code": "Product_Code",
    "description": "Product_Description",
    "qty_ordered": "Quantity",
    "unit_price": "Unit_Price",
    "net_amount": "Net_Amount",
    "tax_amount": "VAT_Amount",
    "gross_amount": "Total_Amount",
}

SAGE_ORDER_OUTPUT_ORDER = [
    "Order_Number", "Customer_ID", "Order_Date", "PO_Reference",
    "Product_Code", "Product_Description", "Quantity", "Unit_Price",
    "Net_Amount", "VAT_Amount", "Total_Amount", "Status"
]

SAGE_PRODUCT_MAP = {
    "stock_code": "Product_Code",
    "description": "Product_Description",
    "category": "Category",
    "sales_price": "Sales_Price",
    "unit_of_sale": "Unit",
    "qty_in_stock": "Stock_Level",
    "reorder_level": "Reorder_Level",
}

SAGE_PRODUCT_OUTPUT_ORDER = [
    "Product_Code", "Product_Description", "Category", "Unit",
    "Sales_Price", "Stock_Level", "Reorder_Level", "Category_Type"
]


# ═══════════════════════════════════════════════════════════════
# Processing Logic
# ═══════════════════════════════════════════════════════════════

def detect_sage_file_type(headers):
    """Detect Sage file type based on column header signatures."""
    hdrs = [h.strip().lower().replace(" ", "_") for h in headers]
    
    # Check Customer signature
    cust_keys = {"account_ref", "name", "telephone", "postcode"}
    if len(cust_keys.intersection(hdrs)) >= 2 and "stock_code" not in hdrs:
        return "Customer"
    
    # Check Sales Order signature
    order_keys = {"order_number", "cust_order_number", "qty_ordered", "net_amount"}
    if len(order_keys.intersection(hdrs)) >= 2 or ("order_number" in hdrs and "stock_code" in hdrs):
        return "Sales Order"
    
    # Check Products signature
    prod_keys = {"stock_code", "sales_price", "qty_in_stock", "unit_of_sale"}
    if len(prod_keys.intersection(hdrs)) >= 2:
        return "Product"
    
    return "Generic CSV"


def process_sage_customer_csv(input_path, output_path, segment_val="Standard", status_val="Active"):
    """Process Sage Customer CSV: clean headers, format phones, reorder columns."""
    infile, enc = open_csv_file(input_path)
    written = 0
    total = 0

    with infile:
        reader = csv.DictReader(infile)
        with open(output_path, "w", newline="", encoding="utf-8") as outfile:
            writer = csv.DictWriter(outfile, fieldnames=SAGE_CUSTOMER_OUTPUT_ORDER)
            writer.writeheader()

            for row in reader:
                total += 1
                norm_row = {k.strip().lower().replace(" ", "_"): v.strip() if v else "" for k, v in row.items()}
                
                new_row = {col: "" for col in SAGE_CUSTOMER_OUTPUT_ORDER}
                for orig_key, new_key in SAGE_CUSTOMER_MAP.items():
                    if new_key in new_row:
                        new_row[new_key] = norm_row.get(orig_key, "")

                # Clean phone fields
                new_row["Phone"] = clean_phone_number(new_row["Phone"])
                new_row["Phone_2"] = clean_phone_number(new_row["Phone_2"])

                # Set custom metadata
                new_row["Segment"] = segment_val
                new_row["Status"] = status_val

                writer.writerow(new_row)
                written += 1

    return total, written


def process_sage_order_csv(input_path, output_path, default_status="Completed"):
    """Process Sage Sales Orders CSV: reorder columns, format dates & amounts."""
    infile, enc = open_csv_file(input_path)
    written = 0
    total = 0

    with infile:
        reader = csv.DictReader(infile)
        with open(output_path, "w", newline="", encoding="utf-8") as outfile:
            writer = csv.DictWriter(outfile, fieldnames=SAGE_ORDER_OUTPUT_ORDER)
            writer.writeheader()

            for row in reader:
                total += 1
                norm_row = {k.strip().lower().replace(" ", "_"): v.strip() if v else "" for k, v in row.items()}

                new_row = {col: "" for col in SAGE_ORDER_OUTPUT_ORDER}
                for orig_key, new_key in SAGE_ORDER_MAP.items():
                    if new_key in new_row:
                        new_row[new_key] = norm_row.get(orig_key, "")

                # Format Order Date
                new_row["Order_Date"] = format_date(new_row["Order_Date"])
                new_row["Status"] = default_status

                writer.writerow(new_row)
                written += 1

    return total, written


def process_sage_product_csv(input_path, output_path, default_type="General"):
    """Process Sage Products CSV: reorder columns, format pricing."""
    infile, enc = open_csv_file(input_path)
    written = 0
    total = 0

    with infile:
        reader = csv.DictReader(infile)
        with open(output_path, "w", newline="", encoding="utf-8") as outfile:
            writer = csv.DictWriter(outfile, fieldnames=SAGE_PRODUCT_OUTPUT_ORDER)
            writer.writeheader()

            for row in reader:
                total += 1
                norm_row = {k.strip().lower().replace(" ", "_"): v.strip() if v else "" for k, v in row.items()}

                new_row = {col: "" for col in SAGE_PRODUCT_OUTPUT_ORDER}
                for orig_key, new_key in SAGE_PRODUCT_MAP.items():
                    if new_key in new_row:
                        new_row[new_key] = norm_row.get(orig_key, "")

                new_row["Category_Type"] = default_type
                writer.writerow(new_row)
                written += 1

    return total, written


def process_sage_split_join(cust_path, order_path, output_dir, name_prefix):
    """Perform join and segmentation between Sage Customers & Orders."""
    cust_file, _ = open_csv_file(cust_path)
    order_file, _ = open_csv_file(order_path)

    customers = {}
    with cust_file:
        c_reader = csv.DictReader(cust_file)
        for row in c_reader:
            norm = {k.strip().lower().replace(" ", "_"): v.strip() if v else "" for k, v in row.items()}
            cid = norm.get("account_ref", "") or norm.get("customer_id", "")
            if cid:
                customers[cid] = norm

    orders = []
    with order_file:
        o_reader = csv.DictReader(order_file)
        for row in o_reader:
            norm = {k.strip().lower().replace(" ", "_"): v.strip() if v else "" for k, v in row.items()}
            orders.append(norm)

    # Active Customers (have orders) vs Inactive
    active_cids = {o.get("account_ref", "") or o.get("customer_id", "") for o in orders if o.get("account_ref") or o.get("customer_id")}
    
    active_cust_rows = []
    inactive_cust_rows = []
    for cid, cdata in customers.items():
        row_out = {
            "Customer_ID": cid,
            "Customer_Name": cdata.get("name", "") or cdata.get("customer_name", ""),
            "Phone": clean_phone_number(cdata.get("telephone", "") or cdata.get("phone", "")),
            "Email": cdata.get("e_mail", "") or cdata.get("email", ""),
            "Postcode": cdata.get("postcode", ""),
            "Balance": cdata.get("balance", "") or cdata.get("current_balance", "0"),
        }
        if cid in active_cids:
            active_cust_rows.append(row_out)
        else:
            inactive_cust_rows.append(row_out)

    fieldnames = ["Customer_ID", "Customer_Name", "Phone", "Email", "Postcode", "Balance"]
    date_str = datetime.now().strftime("%Y-%m-%d")

    active_out = os.path.join(output_dir, f"{name_prefix}_Active_Customers_{date_str}.csv")
    inactive_out = os.path.join(output_dir, f"{name_prefix}_Inactive_Customers_{date_str}.csv")

    with open(active_out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(active_cust_rows)

    with open(inactive_out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(inactive_cust_rows)

    return {
        "total_customers": len(customers),
        "total_orders": len(orders),
        "active_customers": len(active_cust_rows),
        "inactive_customers": len(inactive_cust_rows),
        "active_file": active_out,
        "inactive_file": inactive_out
    }


# ═══════════════════════════════════════════════════════════════
# Desktop GUI Application (Tkinter + Modern Dark Theme)
# ═══════════════════════════════════════════════════════════════

class SageFileOrganiserApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Sage File Organiser")
        self.geometry("980replace_with_720")
        self.geometry("980x720")
        self.minsize(850, 600)

        self._init_theme()
        self._create_layout()

    def _init_theme(self):
        """Configure dark theme colors and ttk styles."""
        self.colors = {
            "bg": "#1e1e1e",
            "fg": "#cccccc",
            "panel": "#252526",
            "panel_border": "#333333",
            "accent": "#007acc",
            "accent_hover": "#1f8ad2",
            "success": "#4ec9b0",
            "warning": "#ce9178",
            "error": "#f14c4c",
            "entry_bg": "#3c3c3c",
            "entry_fg": "#ffffff",
        }

        self.configure(bg=self.colors["bg"])
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=self.colors["bg"], foreground=self.colors["fg"], font=("Segoe UI", 10))
        style.configure("TNotebook", background=self.colors["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=self.colors["panel"], foreground=self.colors["fg"], padding=[16, 8], borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", self.colors["accent"])], foreground=[("selected", "#ffffff")])

        style.configure("TFrame", background=self.colors["bg"])
        style.configure("Panel.TFrame", background=self.colors["panel"], relief="solid", borderwidth=1)

        style.configure("TLabel", background=self.colors["bg"], foreground=self.colors["fg"])
        style.configure("Header.TLabel", font=("Segoe UI", 12, "bold"), foreground="#ffffff")
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground=self.colors["accent"])

        style.configure("TButton", background=self.colors["accent"], foreground="#ffffff", padding=[12, 6], borderwidth=0, font=("Segoe UI", 10, "bold"))
        style.map("TButton", background=[("active", self.colors["accent_hover"])])

        style.configure("Secondary.TButton", background="#3a3d41", foreground="#ffffff")
        style.map("Secondary.TButton", background=[("active", "#4e5258")])

        style.configure("TEntry", fieldbackground=self.colors["entry_bg"], foreground=self.colors["entry_fg"], borderwidth=1)

    def _create_layout(self):
        # Header Banner
        top_frame = ttk.Frame(self, padding=(20, 15, 20, 10))
        top_frame.pack(fill="x")

        title_lbl = ttk.Label(top_frame, text="Sage File Organiser", style="Title.TLabel")
        title_lbl.pack(side="left")

        subtitle_lbl = ttk.Label(top_frame, text="Re-order, clean, transform & split Sage CSV exports", style="TLabel")
        subtitle_lbl.pack(side="left", padx=(15, 0), pady=(4, 0))

        # Notebook / Tabs
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        # Tab 1: Single Sage File Processor
        self.tab1 = ttk.Frame(notebook, padding=15)
        notebook.add(self.tab1, text="Sage File Organizer")
        self._build_tab1()

        # Tab 2: Sage Splitter & Joiner
        self.tab2 = ttk.Frame(notebook, padding=15)
        notebook.add(self.tab2, text="Data Splitter & Joiner")
        self._build_tab2()

    def _build_tab1(self):
        # File selector frame
        panel = ttk.Frame(self.tab1, style="Panel.TFrame", padding=15)
        panel.pack(fill="x", pady=(0, 15))

        lbl = ttk.Label(panel, text="Select Sage CSV File:", style="Header.TLabel")
        lbl.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        self.t1_file_path = tk.StringVar()
        entry = ttk.Entry(panel, textvariable=self.t1_file_path)
        entry.grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 10))

        btn_browse = ttk.Button(panel, text="Browse...", style="Secondary.TButton", command=self._browse_t1_file)
        btn_browse.grid(row=1, column=2)

        # Options
        opts_frame = ttk.Frame(panel, padding=(0, 15, 0, 0))
        opts_frame.grid(row=2, column=0, columnspan=3, sticky="ew")

        ttk.Label(opts_frame, text="Detected Type:").grid(row=0, column=0, sticky="w")
        self.t1_detected_type = tk.StringVar(value="Auto-Detecting...")
        ttk.Label(opts_frame, textvariable=self.t1_detected_type, font=("Segoe UI", 10, "bold"), foreground=self.colors["success"]).grid(row=0, column=1, sticky="w", padx=(5, 20))

        ttk.Label(opts_frame, text="Segment Tag:").grid(row=0, column=2, sticky="w")
        self.t1_segment = tk.StringVar(value="Standard")
        ttk.Entry(opts_frame, textvariable=self.t1_segment, width=15).grid(row=0, column=3, sticky="w", padx=(5, 20))

        ttk.Label(opts_frame, text="Output Directory:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.t1_out_dir = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        ttk.Entry(opts_frame, textvariable=self.t1_out_dir, width=40).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(5, 10), pady=(10, 0))
        ttk.Button(opts_frame, text="Change Folder", style="Secondary.TButton", command=self._browse_t1_dir).grid(row=1, column=3, sticky="w", pady=(10, 0))

        panel.columnconfigure(0, weight=1)

        # Process Button
        btn_process = ttk.Button(self.tab1, text="Process Sage CSV File", command=self._process_t1)
        btn_process.pack(fill="x", pady=(0, 15))

        # Output Log Box
        self.t1_log = self._create_log_box(self.tab1)

    def _build_tab2(self):
        panel = ttk.Frame(self.tab2, style="Panel.TFrame", padding=15)
        panel.pack(fill="x", pady=(0, 15))

        ttk.Label(panel, text="Sage Customer & Sales Order Splitter/Joiner", style="Header.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        ttk.Label(panel, text="Customer CSV:").grid(row=1, column=0, sticky="w")
        self.t2_cust_path = tk.StringVar()
        ttk.Entry(panel, textvariable=self.t2_cust_path).grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        ttk.Button(panel, text="Browse...", style="Secondary.TButton", command=lambda: self._browse_file(self.t2_cust_path)).grid(row=1, column=2, pady=5)

        ttk.Label(panel, text="Sales Order CSV:").grid(row=2, column=0, sticky="w")
        self.t2_order_path = tk.StringVar()
        ttk.Entry(panel, textvariable=self.t2_order_path).grid(row=2, column=1, sticky="ew", padx=10, pady=5)
        ttk.Button(panel, text="Browse...", style="Secondary.TButton", command=lambda: self._browse_file(self.t2_order_path)).grid(row=2, column=2, pady=5)

        ttk.Label(panel, text="Output File Prefix:").grid(row=3, column=0, sticky="w", pady=5)
        self.t2_prefix = tk.StringVar(value="SageExport")
        ttk.Entry(panel, textvariable=self.t2_prefix, width=20).grid(row=3, column=1, sticky="w", padx=10, pady=5)

        ttk.Label(panel, text="Output Directory:").grid(row=4, column=0, sticky="w", pady=5)
        self.t2_out_dir = tk.StringVar(value=os.path.expanduser("~/Downloads"))
        ttk.Entry(panel, textvariable=self.t2_out_dir).grid(row=4, column=1, sticky="ew", padx=10, pady=5)
        ttk.Button(panel, text="Browse...", style="Secondary.TButton", command=lambda: self._browse_dir(self.t2_out_dir)).grid(row=4, column=2, pady=5)

        panel.columnconfigure(1, weight=1)

        btn_process = ttk.Button(self.tab2, text="Join & Segment Sage Files", command=self._process_t2)
        btn_process.pack(fill="x", pady=(0, 15))

        self.t2_log = self._create_log_box(self.tab2)

    def _create_log_box(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)

        txt = tk.Text(frame, bg="#101010", fg="#cccccc", font=("Consolas", 10), borderwidth=0, wrap="word")
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=scrollbar.set)

        txt.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Log color tags
        txt.tag_config("header", foreground="#569cd6", font=("Consolas", 10, "bold"))
        txt.tag_config("info", foreground="#cccccc")
        txt.tag_config("success", foreground="#4ec9b0", font=("Consolas", 10, "bold"))
        txt.tag_config("warning", foreground="#ce9178")
        txt.tag_config("error", foreground="#f14c4c", font=("Consolas", 10, "bold"))

        return txt

    def _log(self, text_widget, msg, tag="info"):
        text_widget.insert("end", msg + "\n", tag)
        text_widget.see("end")

    def _browse_t1_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if path:
            self.t1_file_path.set(path)
            try:
                infile, enc = open_csv_file(path)
                with infile:
                    reader = csv.reader(infile)
                    headers = next(reader, [])
                    file_type = detect_sage_file_type(headers)
                    self.t1_detected_type.set(file_type)
                    self._log(self.t1_log, f"File loaded: {os.path.basename(path)} (Encoding: {enc})", "info")
                    self._log(self.t1_log, f"Auto-Detected File Type: {file_type}", "header")
            except Exception as e:
                self.t1_detected_type.set("Unknown")
                self._log(self.t1_log, f"Error reading file: {e}", "error")

    def _browse_t1_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.t1_out_dir.set(d)

    def _browse_file(self, var):
        p = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if p:
            var.set(p)

    def _browse_dir(self, var):
        d = filedialog.askdirectory()
        if d:
            var.set(d)

    def _process_t1(self):
        path = self.t1_file_path.get().strip()
        out_dir = self.t1_out_dir.get().strip()
        segment = self.t1_segment.get().strip()

        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Please select a valid Sage CSV file.")
            return
        if not out_dir or not os.path.exists(out_dir):
            messagebox.showerror("Error", "Please select a valid output directory.")
            return

        def run():
            try:
                self._log(self.t1_log, "\n" + "="*50, "header")
                self._log(self.t1_log, "Starting Sage CSV processing...", "header")

                infile, enc = open_csv_file(path)
                with infile:
                    reader = csv.reader(infile)
                    headers = next(reader, [])
                
                ftype = detect_sage_file_type(headers)
                base_name = os.path.splitext(os.path.basename(path))[0]
                date_str = datetime.now().strftime("%Y-%m-%d")
                out_name = f"Cleaned_{base_name}_{date_str}.csv"
                out_path = os.path.join(out_dir, out_name)

                if ftype == "Customer":
                    tot, wrt = process_sage_customer_csv(path, out_path, segment)
                elif ftype == "Sales Order":
                    tot, wrt = process_sage_order_csv(path, out_path)
                elif ftype == "Product":
                    tot, wrt = process_sage_product_csv(path, out_path)
                else:
                    tot, wrt = process_sage_customer_csv(path, out_path, segment)

                self._log(self.t1_log, f"✅ Processing Complete!", "success")
                self._log(self.t1_log, f"   Detected File Type: {ftype}", "info")
                self._log(self.t1_log, f"   Total Input Rows:   {tot:,}", "info")
                self._log(self.t1_log, f"   Rows Written:       {wrt:,}", "success")
                self._log(self.t1_log, f"   Output Saved To:    {out_path}", "success")
            except Exception as e:
                self._log(self.t1_log, f"❌ Error: {e}", "error")

        threading.Thread(target=run, daemon=True).start()

    def _process_t2(self):
        cpath = self.t2_cust_path.get().strip()
        opath = self.t2_order_path.get().strip()
        out_dir = self.t2_out_dir.get().strip()
        prefix = self.t2_prefix.get().strip() or "SageExport"

        if not cpath or not opath:
            messagebox.showerror("Error", "Please select both Customer CSV and Sales Order CSV files.")
            return

        def run():
            try:
                self._log(self.t2_log, "\n" + "="*50, "header")
                self._log(self.t2_log, "Starting Sage Splitter & Joiner...", "header")
                
                stats = process_sage_split_join(cpath, opath, out_dir, prefix)

                self._log(self.t2_log, f"✅ Segmentation Complete!", "success")
                self._log(self.t2_log, f"   Total Customers Loaded: {stats['total_customers']:,}", "info")
                self._log(self.t2_log, f"   Total Orders Loaded:    {stats['total_orders']:,}", "info")
                self._log(self.t2_log, f"   Active Customers:       {stats['active_customers']:,}", "success")
                self._log(self.t2_log, f"   Inactive Customers:     {stats['inactive_customers']:,}", "warning")
                self._log(self.t2_log, f"   Saved Active CSV:       {stats['active_file']}", "success")
                self._log(self.t2_log, f"   Saved Inactive CSV:     {stats['inactive_file']}", "success")
            except Exception as e:
                self._log(self.t2_log, f"❌ Error: {e}", "error")

        threading.Thread(target=run, daemon=True).start()


if __name__ == "__main__":
    app = SageFileOrganiserApp()
    app.mainloop()
