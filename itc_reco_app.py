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
load_dotenv(os.path.join(get_app_dir(), ".env"))

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

        self.title("ITC Reconciliation Tool")
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
        self.l1.bind("<Button-1>", lambda e: self.toggle_type_filter("2B"))
        self.lbl_l1.bind("<Button-1>", lambda e: self.toggle_type_filter("2B"))

        self.l2 = tk.Frame(legend_frame, bg=BG_COLOR, cursor="hand2")
        self.l2.pack(fill="x", pady=2)
        self.can_l2 = tk.Canvas(self.l2, width=12, height=12, bg=ACCENT_BOOKS, highlightthickness=0)
        self.can_l2.pack(side="left", padx=(0, 5))
        self.lbl_l2 = ttk.Label(self.l2, text="Booking Done / 2B Pending", font=("Arial", 8))
        self.lbl_l2.pack(side="left")
        self.l2.bind("<Button-1>", lambda e: self.toggle_type_filter("BOOKS"))
        self.lbl_l2.bind("<Button-1>", lambda e: self.toggle_type_filter("BOOKS"))

        self.status_label = ttk.Label(self.sidebar, text="Ready", foreground=STALE_COLOR)
        self.status_label.pack(side="bottom", pady=10)

        # Main Table Area
        right_frame = ttk.Frame(self.content_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        ttk.Label(right_frame, text="Mismatched Records Table", font=("Arial", 12, "bold")).pack(anchor="w", pady=(0, 10))

        # Table Container with Scrollbar
        table_container = ttk.Frame(right_frame)
        table_container.pack(fill="both", expand=True)
        
        self.tree = ttk.Treeview(table_container, columns=("sn", "date", "due_date", "type", "invoice", "tax_val"), show="headings")
        
        self.tree.heading("sn", text="Supplier Name")
        self.tree.heading("date", text="Invoice Date")
        self.tree.heading("due_date", text="Due Date")
        self.tree.heading("type", text="Mismatch Type")
        self.tree.heading("invoice", text="Invoice Number")
        self.tree.heading("tax_val", text="Taxable Value")
        
        self.tree.column("sn", width=250, minwidth=150)
        self.tree.column("date", width=100, minwidth=100, anchor="center")
        self.tree.column("due_date", width=100, minwidth=100, anchor="center")
        self.tree.column("type", width=200, minwidth=150)
        self.tree.column("invoice", width=120, minwidth=100, anchor="center")
        self.tree.column("tax_val", width=110, minwidth=100, anchor="e")
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        v_scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=v_scroll.set)
        
        self.tree.pack(side="left", fill="both", expand=True)

    def toggle_type_filter(self, type_code):
        if type_code in self.active_types:
            self.active_types.remove(type_code)
        else:
            self.active_types.add(type_code)
        
        # UI Feedback for toggles
        self.can_l1.configure(bg=ACCENT_2B if "2B" in self.active_types else "#e0e0e0")
        self.can_l2.configure(bg=ACCENT_BOOKS if "BOOKS" in self.active_types else "#e0e0e0")
        
        self.apply_filter()

    def setup_footer(self):
        footer_frame = ttk.Frame(self)
        footer_frame.pack(side="bottom", fill="x", padx=20, pady=10)
        
        ttk.Label(footer_frame, text="© Nagarkot Forwarders Pvt Ltd", font=("Arial", 8), foreground=STALE_COLOR).pack(side="left")

        # Re-using ModernButton for Footer as well
        self.gen_btn_frame = tk.Frame(footer_frame, bg=SUCCESS_COLOR)
        self.gen_btn_frame.pack(side="right")
        self.generate_btn = ModernButton(self.gen_btn_frame, text="Generate Sheet1 Output", command=self.generate_output, color=SUCCESS_COLOR)
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
            target_sheet = None
            
            # Prioritize 25-26 if current year
            for sheet in xl.sheet_names:
                if "2B RECO 25-26" in sheet.upper():
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                # Fallback to general search
                for sheet in xl.sheet_names:
                    if "2B RECO" in sheet.upper():
                        target_sheet = sheet
                        break
            
            if not target_sheet:
                messagebox.showerror("Error", "Could not find '2B RECO' sheet in file.")
                return

            # Read Excel without assuming headers are on line 0 (sometimes they are deeper)
            df = pd.read_excel(path, sheet_name=target_sheet)
            
            # Clean column names (strip spaces, etc)
            df.columns = [str(c).strip() for c in df.columns]

            if 'Status' not in df.columns:
                messagebox.showerror("Error", f"Sheet '{target_sheet}' is missing the 'Status' column.\n\nAvailable columns: {list(df.columns)}")
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
            
        if 'Taxable Value' in data.columns:
            # Clean Taxable Value: use .loc to avoid SettingWithCopyWarning
            data = data.copy()
            data['TaxVal_num'] = pd.to_numeric(data['Taxable Value'], errors='coerce')
            data = data.sort_values(by='TaxVal_num', ascending=False, na_position='last')
            # Drop the temporary column
            data = data.drop(columns=['TaxVal_num'])
            
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
            
            type_text = "2B Done / Book Pending" if origin == "2B" else "Booking Done / 2B Pending"
            
            self.tree.insert("", "end", values=(supplier, inv_date, current_due_date, type_text, inv_no, tax_val))
            
        self.status_label.configure(text=f"Showing {len(self.filtered_data)} records", foreground=THEME_COLOR)

    def get_access_token(self):
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
            response = requests.post(url, params=params)
            data = response.json()
            return data.get("access_token")
        except Exception as e:
            print(f"Token Error: {e}")
            return None

    def push_to_zoho(self):
        if self.filtered_data is None or self.filtered_data.empty:
            messagebox.showwarning("Warning", "No data to push.")
            return

        selected_items = self.tree.selection()
        
        if selected_items:
            data_to_push = []
            for item_id in selected_items:
                item_values = self.tree.item(item_id, 'values')
                inv_no = item_values[3]
                supplier = item_values[0]
                match = self.filtered_data[(self.filtered_data['Invoice No.'].astype(str) == str(inv_no)) & 
                                         (self.filtered_data['Supplier Name'] == supplier)]
                if not match.empty:
                    data_to_push.append(match.iloc[0])
            
            records_to_push = pd.DataFrame(data_to_push)
            confirm_msg = f"Push {len(records_to_push)} SELECTED records to Shakti?"
        else:
            records_to_push = self.filtered_data
            confirm_msg = f"Push ALL {len(records_to_push)} filtered records to Shakti?"

        refresh_token = os.getenv("ZOHO_REFRESH_TOKEN")
        if not refresh_token:
            messagebox.showerror("Error", "Zoho Refresh Token missing in .env file.")
            return

        if not messagebox.askyesno("Confirm", confirm_msg):
            return

        self.zoho_btn.configure_state("disabled")
        self.status_label.configure(text="Pushing to Shakti...", foreground=THEME_COLOR)
        
        threading.Thread(target=self._push_task, args=(records_to_push,), daemon=True).start()

    def _push_task(self, records):
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

        # Batching logic: Zoho Creator V2 supports bulk record creation (max 100 per request)
        batch_size = 100
        success_count = 0
        error_count = 0
        
        # Convert records to list of dicts for easier batching
        all_records_data = []
        for _, row in records.iterrows():
            origin = str(row.get('Origin', '')).strip().upper()
            mismatch_type = "2B Done / Book Pending" if origin == "2B" else "Booking Done / 2B Pending"
            
            all_records_data.append({
                "Supplier_Name": str(row.get('Supplier Name', '')),
                "Invoice_Date": self.format_date_str(row.get('Invoice Date')),
                "Due_Date": self.get_formatted_due_date(),
                "Mismatch_Type": mismatch_type,
                "Invoice_No": str(row.get('Invoice No.', '')),
                "Taxable_Value": str(row.get('Taxable Value', '0'))
            })

        # Process in batches
        for i in range(0, len(all_records_data), batch_size):
            batch = all_records_data[i:i + batch_size]
            payload = {"data": batch}
            
            # Retry Logic for the batch
            max_retries = 3
            attempt = 0
            pushed = False
            
            while attempt < max_retries and not pushed:
                try:
                    resp = requests.post(url, json=payload, headers=headers, timeout=30)
                    
                    if resp.status_code in [200, 201]:
                        # Handle both list and dict response formats from Zoho
                        resp_data = resp.json()
                        batch_results = []
                        
                        if isinstance(resp_data, list):
                            batch_results = resp_data
                        elif isinstance(resp_data, dict):
                            # Try result -> data first
                            result_obj = resp_data.get("result", {})
                            if isinstance(result_obj, dict):
                                batch_results = result_obj.get("data", [])
                            elif isinstance(result_obj, list):
                                batch_results = result_obj
                            
                            # Fallback: if no individual results yet, check the overall code
                            if not batch_results and resp_data.get("code") == 3000:
                                success_count += len(batch)
                                pushed = True
                                continue

                        if not batch_results:
                            print(f"Unexpected response format: {resp_data}")
                            error_count += len(batch)
                        else:
                            for res in batch_results:
                                # Ensure 'res' is a dictionary before calling .get()
                                if isinstance(res, dict):
                                    code = str(res.get("code", ""))
                                    status = str(res.get("status", ""))
                                    if code == "3000" or status == "success":
                                        success_count += 1
                                    else:
                                        error_count += 1
                                        err_msg = res.get('message') or res.get('error') or 'Record error'
                                        print(f"Record Error: {err_msg}")
                                else:
                                    # If it's not a dict, we can't parse it easily
                                    print(f"Non-dict record result: {res}")
                                    error_count += 1
                        pushed = True
                    elif resp.status_code == 429: # Rate limit
                        time.sleep(5 * (attempt + 1))
                    else:
                        print(f"!!! Batch Failed !!!")
                        print(f"Status Code: {resp.status_code}")
                        print(f"Response: {resp.text}")
                        if resp.status_code >= 500:
                            time.sleep(2)
                        else:
                            break
                except Exception as e:
                    print(f"Batch Request Attempt {attempt+1} Exception: {str(e)}")
                    time.sleep(2)
                attempt += 1

            if not pushed:
                error_count += len(batch)
            
            # Small delay between batches to stay safe with rate limits
            time.sleep(0.5)

        self.after(0, lambda: self._finalize_push(success_count, error_count))

    def _finalize_push(self, success, error):
        self.zoho_btn.configure_state("normal")
        self.status_label.configure(text=f"Push Complete: {success} Success, {error} Failed", 
                                  foreground=SUCCESS_COLOR if error == 0 else "red")
        
        msg = f"Data push complete.\nSuccess: {success}"
        if error > 0:
            msg += f"\nFailed: {error}"
        
        messagebox.showinfo("Zoho Push", msg)

    def generate_output(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path: return

        try:
            output_rows = []
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
                    'Invoice not booked/ITC not Claimed': inv_no if origin == '2B' else '',
                    'ITC not appeared in 2B-Follow up': inv_no if origin == 'BOOKS' else '',
                    'POC': ''
                })

            final_df = pd.DataFrame(output_rows)
            cols = ['Supplier Name', 'Invoice Date', 'Invoice not booked/ITC not Claimed', 'ITC not appeared in 2B-Follow up', 'POC']
            final_df = final_df[cols]

            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            
            messagebox.showinfo("Success", f"Output generated successfully at:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")

if __name__ == "__main__":
    app = ITCRecoApp()
    app.mainloop()
