import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from PIL import Image, ImageTk
import os
import sys
import threading
from datetime import datetime
import requests
import json
import time
import calendar
from tkcalendar import DateEntry
from dotenv import load_dotenv

def get_base_path():
    """Get path for internal bundled assets."""
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_app_dir():
    """Get path for external writable config files."""
    if hasattr(sys, 'frozen'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

# Load configuration
load_dotenv(os.path.join(get_base_path(), ".env"))

# --- Constants ---
THEME_COLOR = "#0056b3"  # Nagarkot Blue
BG_COLOR = "#ffffff"     # White
TEXT_COLOR = "#000000"
SECONDARY_BG = "#f8f9fa"
STALE_COLOR = "#6c757d"
SUCCESS_COLOR = "#28a745"
ACCENT_2B = "#D84315"
ACCENT_BOOKS = "#C2185B"

def get_asset_path(filename):
    return os.path.join(get_base_path(), "assets", filename)

class ModernButton(tk.Frame):
    def __init__(self, master, text, command, color=THEME_COLOR, **kwargs):
        super().__init__(master, bg=color, **kwargs)
        self.command = command
        self.color = color
        
        self.label = tk.Label(self, text=text, fg="white", bg=color, font=("Arial", 10, "bold"), pady=8, cursor="hand2")
        self.label.pack(fill="both", expand=True)
        
        self.label.bind("<Button-1>", lambda e: self.command())
        self.label.bind("<Enter>", self.on_enter)
        self.label.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self.configure(bg="#004494")
        self.label.configure(bg="#004494")

    def on_leave(self, e):
        self.configure(bg=self.color)
        self.label.configure(bg=self.color)

    def configure_state(self, state):
        if state == "disabled":
            self.label.unbind("<Button-1>")
            self.label.configure(fg="#aaaaaa", cursor="arrow")
        else:
            self.label.bind("<Button-1>", lambda e: self.command())
            self.label.configure(fg="white", cursor="hand2")

class ITCRecoApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("ITC Reconciliation Tracker")
        self.geometry("1400x850")
        self.state('zoomed')
        self.configure(bg=BG_COLOR)
        
        # Style Configuration
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Treeview", background=BG_COLOR, font=("Arial", 9), rowheight=30)
        self.style.configure("Treeview.Heading", font=("Arial", 10, "bold"), relief="flat", background=SECONDARY_BG)
        self.style.map("Treeview", background=[('selected', THEME_COLOR)], foreground=[('selected', 'white')])
        
        self.style.configure("TFrame", background=BG_COLOR)
        self.style.configure("TLabel", background=BG_COLOR, font=("Arial", 10), foreground=TEXT_COLOR)
        self.style.configure("Header.TLabel", font=("Arial", 18, "bold"), foreground=THEME_COLOR, background=BG_COLOR)
        self.style.configure("SubHeader.TLabel", font=("Arial", 10), foreground=STALE_COLOR, background=BG_COLOR)
        
        self.setup_header()
        
        self.content_frame = ttk.Frame(self)
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.mismatched_data = None
        self.filtered_data = None
        
        # Calculate default due date (last day of current month)
        now = datetime.now()
        last_day = calendar.monthrange(now.year, now.month)[1]
        self.default_due_date = f"{last_day}-{now.strftime('%b-%Y')}"
        
        self.setup_ui()
        self.setup_footer()

    def setup_header(self):
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=20, pady=10)
        
        upper_container = tk.Frame(header_frame, bg=BG_COLOR, height=70)
        upper_container.pack(fill="x")
        upper_container.pack_propagate(False)

        try:
            logo_path = get_asset_path("logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                h_size = 45
                w_size = int((h_size / float(pil_img.size[1])) * float(pil_img.size[0]))
                pil_img = pil_img.resize((w_size, h_size), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(pil_img)
                tk.Label(upper_container, image=self.logo_img, bg=BG_COLOR).pack(side="left", anchor="nw")
        except Exception as e:
            print(f"Logo Error: {e}")
        
        title_frame = ttk.Frame(upper_container)
        title_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        ttk.Label(title_frame, text="NAGARKOT FORWARDERS", style="Header.TLabel").pack(anchor="center")
        ttk.Label(title_frame, text="ITC Reconciliation Tracker", style="SubHeader.TLabel").pack(anchor="center")

    def setup_ui(self):
        # Sidebar
        self.sidebar = ttk.Frame(self.content_frame, width=280, padding=15, relief="solid", borderwidth=1)
        self.sidebar.pack(side="left", fill="y", padx=(0, 20))
        self.sidebar.pack_propagate(False)

        ttk.Label(self.sidebar, text="ITC Tracker", font=("Arial", 14, "bold"), foreground=THEME_COLOR).pack(pady=(10, 30))
        
        self.upload_btn = ModernButton(self.sidebar, text="Upload INPUT File", command=self.upload_file)
        self.upload_btn.pack(pady=10, fill="x")

        # Filter Section
        filter_frame = ttk.Frame(self.sidebar)
        filter_frame.pack(pady=30, fill="x")
        
        ttk.Label(filter_frame, text="Filter by Supplier:", font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 5))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", self.apply_filter)
        self.filter_entry = ttk.Entry(filter_frame, textvariable=self.filter_var, font=("Arial", 10))
        self.filter_entry.pack(fill="x", pady=5)
        
        # Entity Selection Section
        entity_frame = ttk.Frame(self.sidebar)
        entity_frame.pack(pady=10, fill="x")
        ttk.Label(entity_frame, text="Select Entity:", font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 5))
        
        self.entity_var = tk.StringVar(value="")
        self.entity_combo = ttk.Combobox(entity_frame, textvariable=self.entity_var, state="readonly", font=("Arial", 10))
        self.entity_combo['values'] = ("NFPL (HO)", "NLPL", "NFPL (ISD)")
        self.entity_combo.pack(fill="x", pady=5)
        self.entity_combo.bind("<<ComboboxSelected>>", self.on_entity_changed)

        # Due Date Section
        due_date_frame = ttk.Frame(self.sidebar)
        due_date_frame.pack(pady=10, fill="x")
        ttk.Label(due_date_frame, text="Set Due Date (for Shakti):", font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 5))
        
        # Calculate default date object
        now = datetime.now()
        last_day = calendar.monthrange(now.year, now.month)[1]
        default_date = datetime(now.year, now.month, last_day)

        self.due_date_var = tk.StringVar()
        self.due_date_var.trace_add("write", self.sync_due_date_to_tree)
        
        # Use DateEntry for a dropdown calendar
        self.due_date_picker = DateEntry(due_date_frame, 
                                        textvariable=self.due_date_var,
                                        date_pattern='dd-mm-yyyy',
                                        font=("Arial", 10),
                                        background=THEME_COLOR,
                                        foreground='white',
                                        borderwidth=2)
        self.due_date_picker.set_date(default_date)
        self.due_date_picker.pack(fill="x", pady=5)
        
        # Legend Section
        legend_frame = ttk.Frame(self.sidebar)
        legend_frame.pack(side="bottom", fill="x", pady=20)
        
        self.active_types = {"2B", "BOOKS"}
        ttk.Label(legend_frame, text="Type Legend (Toggle on/off):", font=("Arial", 9, "bold")).pack(anchor="w", pady=(0, 5))
        
        self.l1 = tk.Frame(legend_frame, bg=BG_COLOR, cursor="hand2")
        self.l1.pack(fill="x", pady=2)
        self.can_l1 = tk.Canvas(self.l1, width=12, height=12, bg=ACCENT_2B, highlightthickness=0)
        self.can_l1.pack(side="left", padx=(0, 5))
        self.lbl_l1 = ttk.Label(self.l1, text="2B Done / Book Pending", font=("Arial", 8))
        self.lbl_l1.pack(side="left")
        self.l1.bind("<Button-1>", lambda e: self.toggle_legend_1())
        self.lbl_l1.bind("<Button-1>", lambda e: self.toggle_legend_1())

        self.l2 = tk.Frame(legend_frame, bg=BG_COLOR, cursor="hand2")
        self.l2.pack(fill="x", pady=2)
        self.can_l2 = tk.Canvas(self.l2, width=12, height=12, bg=ACCENT_BOOKS, highlightthickness=0)
        self.can_l2.pack(side="left", padx=(0, 5))
        self.lbl_l2 = ttk.Label(self.l2, text="Booking Done / 2B Pending", font=("Arial", 8))
        self.lbl_l2.pack(side="left")
        self.l2.bind("<Button-1>", lambda e: self.toggle_legend_2())
        self.lbl_l2.bind("<Button-1>", lambda e: self.toggle_legend_2())

        self.status_label = ttk.Label(self.sidebar, text="Ready", foreground=STALE_COLOR)
        self.status_label.pack(side="bottom", pady=10)

        # Main Table Area
        right_frame = ttk.Frame(self.content_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        ttk.Label(right_frame, text="Mismatched Records Table", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))

        # Table Container with Scrollbar
        table_container = ttk.Frame(right_frame)
        table_container.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(table_container, columns=("sn", "date", "due_date", "type", "invoice", "tax_val", "status"), show="headings")
        
        self.tree.heading("sn", text="Supplier Name")
        self.tree.heading("date", text="Invoice Date")
        self.tree.heading("due_date", text="Due Date")
        self.tree.heading("type", text="Mismatch Type")
        self.tree.heading("invoice", text="Invoice Number")
        self.tree.heading("tax_val", text="Taxable Value")
        self.tree.heading("status", text="Status")
        
        self.tree.column("sn", width=200, minwidth=150)
        self.tree.column("date", width=100, minwidth=100, anchor="center")
        self.tree.column("due_date", width=100, minwidth=100, anchor="center")
        self.tree.column("type", width=180, minwidth=150)
        self.tree.column("invoice", width=120, minwidth=100, anchor="center")
        self.tree.column("tax_val", width=110, minwidth=100, anchor="e")
        self.tree.column("status", width=250, minwidth=150, anchor="center")
        
        # Tags for status coloring
        self.tree.tag_configure("success", foreground="#15803d")
        self.tree.tag_configure("error", foreground="#dc3545")
        self.tree.tag_configure("checking", foreground=BRAND_BLUE if "BRAND_BLUE" in globals() else THEME_COLOR)
        self.tree.tag_configure("skipped", foreground="#b45309")
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        v_scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=v_scroll.set)
        
        self.tree.pack(side="left", fill="both", expand=True)

        # Context menu for copying key fields from preview rows
        self.tree_context_menu = tk.Menu(self, tearoff=0)
        self.tree_context_menu.add_command(label="Copy Supplier Name", command=lambda: self.copy_tree_context_value("sn"))
        self.tree_context_menu.add_command(label="Copy Invoice Number", command=lambda: self.copy_tree_context_value("invoice"))
        self.tree.bind("<Button-3>", self.show_tree_context_menu)
        self.tree.bind("<Button-2>", self.show_tree_context_menu)

    def toggle_legend_1(self):
        code = "6A" if self.entity_var.get() == "NFPL (ISD)" else "2B"
        self.toggle_type_filter(code)

    def toggle_legend_2(self):
        self.toggle_type_filter("BOOKS")

    def toggle_type_filter(self, type_code):
        if type_code in self.active_types:
            self.active_types.remove(type_code)
        else:
            self.active_types.add(type_code)
        
        self.update_legend_ui()
        self.apply_filter()

    def update_legend_ui(self):
        entity = self.entity_var.get()
        is_isd = (entity == "NFPL (ISD)")
        code_l1 = "6A" if is_isd else "2B"
        code_l2 = "BOOKS"
        
        self.can_l1.configure(bg=ACCENT_2B if code_l1 in self.active_types else "#e0e0e0")
        self.can_l2.configure(bg=ACCENT_BOOKS if code_l2 in self.active_types else "#e0e0e0")

    def on_entity_changed(self, event=None):
        entity = self.entity_var.get()
        if entity == "NFPL (ISD)":
            self.lbl_l1.configure(text="6A Done / Book Pending")
            self.lbl_l2.configure(text="Booking Done / 6A Pending")
            if "2B" in self.active_types:
                self.active_types.remove("2B")
                self.active_types.add("6A")
        else:
            self.lbl_l1.configure(text="2B Done / Book Pending")
            self.lbl_l2.configure(text="Booking Done / 2B Pending")
            if "6A" in self.active_types:
                self.active_types.remove("6A")
                self.active_types.add("2B")
        
        self.update_legend_ui()
        self.apply_filter()

    def setup_footer(self):
        footer_frame = ttk.Frame(self)
        footer_frame.pack(side="bottom", fill="x", padx=20, pady=10)
        
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", font=("Arial", 8), foreground=STALE_COLOR).pack(side="left")

        # Re-using ModernButton for Footer as well
        self.gen_btn_frame = tk.Frame(footer_frame, bg=SUCCESS_COLOR)
        self.gen_btn_frame.pack(side="right")
        self.generate_btn = ModernButton(self.gen_btn_frame, text="Generate Excel Output", command=self.generate_output, color=SUCCESS_COLOR)
        self.generate_btn.pack()
        self.generate_btn.configure_state("disabled")

        self.zoho_btn_frame = tk.Frame(footer_frame, bg=THEME_COLOR)
        self.zoho_btn_frame.pack(side="right", padx=(0, 10))
        self.zoho_btn = ModernButton(self.zoho_btn_frame, text="Push to Shakti", command=self.push_to_zoho, color=THEME_COLOR)
        self.zoho_btn.pack()
        self.zoho_btn.configure_state("disabled")

    def format_date_str(self, date_val):
        if pd.isna(date_val): return ""
        try:
            if isinstance(date_val, str):
                date_val = pd.to_datetime(date_val)
            return date_val.strftime('%d-%b-%Y')
        except:
            return str(date_val)

    def upload_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path: return

        try:
            xl = pd.ExcelFile(path)
            # Default to the first sheet if there are no specific naming requirements
            target_sheet = xl.sheet_names[0]

            # Read Excel without assuming headers are on line 0 (sometimes they are deeper)
            df = pd.read_excel(path, sheet_name=target_sheet)
            
            # Clean column names (strip spaces, etc)
            df.columns = [str(c).strip() for c in df.columns]

            required_cols = ['Supplier Name', 'Invoice Date', 'Invoice No.', 'Taxable Value', 'Status', 'Origin']
            missing_cols = [col for col in required_cols if col not in df.columns]

            if missing_cols:
                err_msg = (
                    f"The selected sheet '{target_sheet}' is missing required columns:\n"
                    f"{', '.join(missing_cols)}\n\n"
                    f"Please ensure your input Excel file contains these exact columns:\n"
                    f"Supplier Name, Invoice Date, Invoice No., Taxable Value, Status, Origin"
                )
                messagebox.showerror("Invalid Excel Format", err_msg)
                return
            
            # Filter logic: Status != 'Matched' AND Remarks-1 is empty
            condition = (df['Status'].astype(str).str.upper() != 'MATCHED')
            
            # Additional check for Remarks-1 if column exists
            if 'Remarks-1' in df.columns:
                condition &= (df['Remarks-1'].isna() | (df['Remarks-1'].astype(str).str.strip() == ""))
            
            self.mismatched_data = df[condition].copy()
            
            if self.mismatched_data.empty:
                messagebox.showinfo("Info", "No mismatched records found.")
                for item in self.tree.get_children(): self.tree.delete(item)
                return

            self.status_label.configure(text=f"Loaded {len(self.mismatched_data)} mismatches", foreground=SUCCESS_COLOR)
            self.apply_filter()
            self.generate_btn.configure_state("normal")
            self.zoho_btn.configure_state("normal")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load: {e}")

    def upload_poc_register(self):
        # Deprecated
        pass

    def get_formatted_due_date(self):
        """Convert the numeric date from picker to dd-MMM-yyyy."""
        raw_date = self.due_date_var.get()
        try:
            # DateEntry uses the pattern 'dd-mm-yyyy'
            dt = datetime.strptime(raw_date, '%d-%m-%Y')
            return dt.strftime('%d-%b-%Y')
        except:
            return raw_date

    def sync_due_date_to_tree(self, *args):
        if not hasattr(self, "tree"): return
        new_due_date = self.get_formatted_due_date()
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            if len(values) > 2:
                values[2] = new_due_date
                self.tree.item(item, values=values)

    def apply_filter(self, *args):
        if self.mismatched_data is None: return
        
        filter_text = self.filter_var.get().strip().upper()
        
        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        data = self.mismatched_data
        
        # Apply Text Filter
        if filter_text:
            data = data[data['Supplier Name'].astype(str).str.upper().str.contains(filter_text)]
            
        # Apply Type Filter (Multi-toggle)
        data = data[data['Origin'].astype(str).str.upper().isin(self.active_types)]
            
        # Sort by Invoice Date descending
        data = data.copy()
        data['Invoice_Date_dt'] = pd.to_datetime(data['Invoice Date'], errors='coerce')
        data = data.sort_values(by='Invoice_Date_dt', ascending=False, na_position='last')
        data = data.drop(columns=['Invoice_Date_dt'])
            
        self.filtered_data = data
        current_due_date = self.get_formatted_due_date()

        for _, row in self.filtered_data.iterrows():
            origin = str(row.get('Origin', '')).strip().upper()
            inv_no = str(row.get('Invoice No.', ''))
            supplier = row.get('Supplier Name', '')
            inv_date = self.format_date_str(row.get('Invoice Date'))
            tax_val = row.get('Taxable Value', 0)
            try:
                tax_val = f"{float(tax_val):,.2f}"
            except:
                tax_val = str(tax_val)
            
            if self.entity_var.get() == "NFPL (ISD)":
                type_text = "6A Done / Book Pending" if origin == "6A" else "Booking Done / 6A Pending"
            else:
                type_text = "2B Done / Book Pending" if origin == "2B" else "Booking Done / 2B Pending"
            status_text = "Pending"
            
            self.tree.insert("", "end", values=(supplier, inv_date, current_due_date, type_text, inv_no, tax_val, status_text))
            
        self.status_label.configure(text=f"Showing {len(self.filtered_data)} records", foreground=THEME_COLOR)

    def show_tree_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self.tree.selection_set(item)
        try:
            self.tree_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.tree_context_menu.grab_release()

    def copy_tree_context_value(self, column_key):
        if not self.tree.selection():
            return
        values = self.tree.item(self.tree.selection()[0], "values")
        col_idx = {"sn": 0, "invoice": 4}.get(column_key)
        if col_idx is None or col_idx >= len(values):
            return
        value = str(values[col_idx]).strip()
        if not value:
            return
        self.clipboard_clear()
        self.clipboard_append(value)
        field_name = "Supplier Name" if column_key == "sn" else "Invoice Number"
        self.status_label.configure(text=f"Copied {field_name} to clipboard", foreground=THEME_COLOR)

    def get_access_token(self):
        """Get Zoho access token, using cached version if available."""
        # Reuse cached token if we already have one for this session
        if hasattr(self, '_cached_access_token') and self._cached_access_token:
            return self._cached_access_token

        client_id = os.getenv("ZOHO_CLIENT_ID")
        client_secret = os.getenv("ZOHO_CLIENT_SECRET")
        refresh_token = os.getenv("ZOHO_REFRESH_TOKEN")
        
        if not all([client_id, client_secret, refresh_token]):
            return None

        url = "https://accounts.zoho.in/oauth/v2/token"
        params = {
            "refresh_token": refresh_token,
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "refresh_token"
        }
        
        try:
            response = requests.post(url, params=params, timeout=15)
            data = response.json()
            token = data.get("access_token")
            if token:
                self._cached_access_token = token
            return token
        except Exception as e:
            print(f"Token Error: {e}")
            return None

    def push_to_zoho(self):
        if self.filtered_data is None or self.filtered_data.empty:
            messagebox.showwarning("Warning", "No data to push.")
            return

        selected_entity = self.entity_var.get()
        if not selected_entity:
            messagebox.showwarning("Warning", "Please select an Entity before pushing to Shakti.")
            return

        selected_items = self.tree.selection()
        
        if selected_items:
            data_to_push = []
            for item_id in selected_items:
                item_values = self.tree.item(item_id, 'values')
                inv_no = item_values[4]  # index 4 is invoice number now
                supplier = item_values[0]
                match = self.filtered_data[(self.filtered_data['Invoice No.'].astype(str) == str(inv_no)) & 
                                         (self.filtered_data['Supplier Name'] == supplier)]
                if not match.empty:
                    data_to_push.append((item_id, match.iloc[0]))
            
            records_to_push_tuples = data_to_push
            confirm_msg = f"Push {len(records_to_push_tuples)} SELECTED records to Shakti?"
        else:
            records_to_push_tuples = []
            for item_id in self.tree.get_children():
                item_values = self.tree.item(item_id, 'values')
                inv_no = item_values[4]
                supplier = item_values[0]
                match = self.filtered_data[(self.filtered_data['Invoice No.'].astype(str) == str(inv_no)) & 
                                         (self.filtered_data['Supplier Name'] == supplier)]
                if not match.empty:
                    records_to_push_tuples.append((item_id, match.iloc[0]))
                    
            confirm_msg = f"Push ALL {len(records_to_push_tuples)} filtered records to Shakti?"

        refresh_token = os.getenv("ZOHO_REFRESH_TOKEN")
        if not refresh_token:
            messagebox.showerror("Error", "Zoho Refresh Token missing in .env file.")
            return

        if not messagebox.askyesno("Confirm", confirm_msg):
            return

        self.zoho_btn.configure_state("disabled")
        self.status_label.configure(text="Checking for duplicates and pushing to Shakti...", foreground=THEME_COLOR)
        
        threading.Thread(target=self._push_task, args=(records_to_push_tuples, selected_entity), daemon=True).start()

    def update_tree_status(self, item_id, status_text, tag=""):
        try:
            values = list(self.tree.item(item_id, "values"))
            if len(values) >= 7:
                values[6] = status_text
                self.tree.item(item_id, values=values, tags=(tag,) if tag else ())
                self.tree.see(item_id)
        except Exception as e:
            pass

    def check_record_exists(self, base_url, app_owner, app_link_name, token, supplier, inv_date, inv_no):
        report_link_name = os.getenv("ZOHO_REPORT_LINK_NAME", "All_ITC")
        url = f"{base_url}/api/v2/{app_owner}/{app_link_name}/report/{report_link_name}"
        
        # Exclude Date from criteria to prevent Zoho date format mismatches.
        criteria = f'(Supplier_Name == "{supplier}" && Invoice_No == "{inv_no}")'
        params = {"criteria": criteria}
        headers = {
            "Authorization": f"Zoho-oauthtoken {token}",
            "Content-Type": "application/json"
        }
        
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=15)
            print(f"GET {url} params: {params} -> Status: {resp.status_code}")
            
            if resp.status_code == 404:
                return False, "Report 'All_ITC' not found (404)"
            
            if resp.status_code == 200:
                data = resp.json()
                records = data.get("data", [])
                
                # Check results
                for rec in records:
                    return True, None
                    
                return False, None
            elif resp.status_code == 401:
                try:
                    err_json = resp.json()
                    if "oauthscope" in err_json.get("message", "").lower() or err_json.get("code") == 1030:
                        return False, "OAuth Scope Error: Your Zoho Token lacks READ permissions (ZohoCreator.report.READ). Please regenerate your token."
                except:
                    pass
                return False, f"Auth Error (401): {resp.text}"
            else:
                err_msg = f"HTTP {resp.status_code}: {resp.text}"
                print(f"GET Error Response: {err_msg}")
                return False, err_msg
                
        except Exception as e:
            err_msg = str(e)
            print(f"Check Exists Exception: {err_msg}")
            return False, err_msg

    def _push_task(self, records_tuples, entity_name):
        token = self.get_access_token()
        if not token:
            self.after(0, lambda: messagebox.showerror("Error", "Failed to get Zoho Access Token."))
            self.after(0, lambda: self.zoho_btn.configure_state("normal"))
            return

        app_owner = os.getenv("ZOHO_OWNER_NAME")
        app_link_name = os.getenv("ZOHO_APP_LINK_NAME")
        form_link_name = os.getenv("ZOHO_FORM_LINK_NAME")
        base_url = os.getenv("ZOHO_BASE_URL", "https://creator.zoho.in")

        url = f"{base_url}/api/v2/{app_owner}/{app_link_name}/form/{form_link_name}"
        headers = {
            "Authorization": f"Zoho-oauthtoken {token}",
            "Content-Type": "application/json"
        }

        batch_size = 100
        success_count = 0
        skipped_count = 0
        error_count = 0
        
        # Filter out duplicates first
        records_to_push = []
        for item_id, row in records_tuples:
            supplier = str(row.get('Supplier Name', '')).strip()
            inv_date = self.format_date_str(row.get('Invoice Date')).strip()
            inv_no = str(row.get('Invoice No.', '')).strip()
            
            self.after(0, lambda i=item_id: self.update_tree_status(i, "Syncing...", "checking"))
            
            exists, err = self.check_record_exists(base_url, app_owner, app_link_name, token, supplier, inv_date, inv_no)
            
            if err:
                self.after(0, lambda i=item_id: self.update_tree_status(i, "Bypassing Check...", "checking"))
                print(f"Bypassing duplicate check for {inv_no} due to API error: {err}")
                exists = False
            elif exists:
                skipped_count += 1
                self.after(0, lambda i=item_id: self.update_tree_status(i, "Skipped (Already Present)", "skipped"))
                continue
            else:
                self.after(0, lambda i=item_id: self.update_tree_status(i, "Queued...", "checking"))
                
            if not exists:
                origin = str(row.get('Origin', '')).strip().upper()
                if entity_name == "NFPL (ISD)":
                    mismatch_type = "6A Done / Book Pending" if origin == "6A" else "Booking Done / 6A Pending"
                else:
                    mismatch_type = "2B Done / Book Pending" if origin == "2B" else "Booking Done / 2B Pending"
                
                records_to_push.append({
                    "item_id": item_id,
                    "payload": {
                        "Entity_Name": entity_name,
                        "Supplier_Name": supplier,
                        "Invoice_Date": inv_date,
                        "Due_Date": self.get_formatted_due_date(),
                        "Mismatch_Type": mismatch_type,
                        "Invoice_No": inv_no,
                        "Taxable_Value": str(row.get('Taxable Value', '0'))
                    }
                })

        # Process records one-by-one (Zoho Creator V2 Form POST requires single object)
        for item in records_to_push:
            item_id = item["item_id"]
            self.after(0, lambda i=item_id: self.update_tree_status(i, "Syncing...", "checking"))
            
            payload = {"data": item["payload"]}
            
            try:
                resp = requests.post(url, json=payload, headers=headers, timeout=60)
                print(f"POST Record {item['payload'].get('Invoice_No')}: Status {resp.status_code}")
                
                if 200 <= resp.status_code < 300:
                    success_count += 1
                    status_msg = "Updated (No Dup Check)" if err else "Updated"
                    self.after(0, lambda i=item_id, msg=status_msg: self.update_tree_status(i, msg, "success"))
                else:
                    print(f"POST failed: {resp.text}")
                    error_count += 1
                    self.after(0, lambda i=item_id: self.update_tree_status(i, f"Not Updated (Err {resp.status_code})", "error"))
                    
                    if resp.status_code == 401:
                        self._cached_access_token = None
            except Exception as e:
                print(f"POST request error: {e}")
                error_count += 1
                self.after(0, lambda i=item_id: self.update_tree_status(i, "Not Updated (Network Err)", "error"))
            
            # Small delay to prevent rate limits
            time.sleep(0.2)

        self.after(0, lambda: self._finalize_push(success_count, error_count, skipped_count))

    def _finalize_push(self, success, error, skipped):
        self.zoho_btn.configure_state("normal")
        self.status_label.configure(text="Push Complete", foreground=SUCCESS_COLOR)
        
        msg = f"Data push complete.\n\nSuccessfully Pushed: {success}\nAlready Present (Skipped): {skipped}\nErrors: {error}"
        messagebox.showinfo("Shakti Push", msg)

    def generate_output(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path: return

        try:
            output_rows = []
            entity = self.entity_var.get()
            is_isd = (entity == "NFPL (ISD)")
            col_not_booked = 'Invoice not booked/ITC not Claimed'
            col_not_appeared = 'ITC not appeared in 6A-Follow up' if is_isd else 'ITC not appeared in 2B-Follow up'
            match_origin = '6A' if is_isd else '2B'

            # We export all mismatched data, not just filtered ones, unless requested otherwise.
            # Usually users want to filter to see, but export everything.
            for _, row in self.mismatched_data.iterrows():
                origin = str(row.get('Origin', '')).strip().upper()
                inv_no = str(row.get('Invoice No.', ''))
                supplier = row.get('Supplier Name', '')
                inv_date = self.format_date_str(row.get('Invoice Date'))
                
                output_rows.append({
                    'Supplier Name': supplier,
                    'Invoice Date': inv_date,
                    col_not_booked: inv_no if origin == match_origin else '',
                    col_not_appeared: inv_no if origin == 'BOOKS' else '',
                    'POC': ''
                })

            final_df = pd.DataFrame(output_rows)
            cols = ['Supplier Name', 'Invoice Date', col_not_booked, col_not_appeared, 'POC']
            final_df = final_df[cols]

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            messagebox.showinfo("Success", f"Output generated successfully at:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")

if __name__ == "__main__":
    app = ITCRecoApp()
    app.mainloop()
