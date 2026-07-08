import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
from offer_engine import export_to_excel, process_mix_match
import os
import sys
import hashlib
import uuid

# ==========================================
# نظام الحماية والتفعيل (Offline License)
# ==========================================
SECRET_SALT = "LOTUS_PHARMA_2026_SUPER_SECRET_KEY" 

def get_machine_id():
    mac = uuid.getnode()
    return hashlib.md5(str(mac).encode()).hexdigest()[:10].upper()

def generate_expected_key(machine_id):
    raw_string = f"{machine_id}_{SECRET_SALT}"
    return hashlib.sha256(raw_string.encode()).hexdigest()[:16].upper()

def check_license():
    machine_id = get_machine_id()
    expected_key = generate_expected_key(machine_id)
    license_file = "lotus_offer.lic"

    if os.path.exists(license_file):
        with open(license_file, "r") as f:
            if f.read().strip() == expected_key:
                return True 

    root = tk.Tk()
    root.withdraw()
    
    msg = f"هذا الجهاز غير مصرح له باستخدام النظام.\n\nرقم جهازك هو: {machine_id}\n\nبرجاء إرسال هذا الرقم للإدارة للحصول على كود التفعيل."
    user_key = simpledialog.askstring("تفعيل ترخيص لوتس", msg, parent=root)

    if user_key and user_key.strip().upper() == expected_key:
        with open(license_file, "w") as f:
            f.write(user_key.strip().upper())
        messagebox.showinfo("نجاح", "تم تفعيل النظام بنجاح! شكراً لك.")
        root.destroy()
        return True
    else:
        messagebox.showerror("خطأ", "كود التفعيل غير صحيح أو تم الإلغاء. سيتم إغلاق النظام.")
        sys.exit()

# تشغيل فحص الترخيص قبل فتح البرنامج
check_license()
# ==========================================


# --- إعدادات الإصدار ---
APP_VERSION = "1.4"
# -----------------------

# إعدادات المظهر
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class LotusOfferGenerator:
    def __init__(self, root):
        self.root = root
        
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.root.title(f"Lotus Pro Claims - Mix & Match Engine | v{APP_VERSION} | {current_date}")
        self.root.geometry("850x850")
        
        # --- تغيير أيقونة البرنامج لتعمل مع ملفات الـ exe ---
        try:
            if hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            else:
                basedir = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(basedir, "icon.ico")
            self.root.iconbitmap(icon_path)
        except Exception:
            pass
        # ----------------------------------------------------

        self.filepath = ctk.StringVar()
        self.filter_by_date_var = ctk.BooleanVar(value=False)
        self.target_discount_var = ctk.StringVar()
        
        self.available_offers = ["Skip", "1+50%", "1+1", "2+1", "10%", "20%"]
        
        self.company_mappings = {}  
        self.company_checkbox_vars = {} 

        self.branches = [
            "All", "حي السفارات - مدينة نصر", "مول المعز - الشيخ زايد", 
            "السوق الشرقي - الرحاب", "الرؤى مول - العاشر من رمضان", 
            "تريومف - مصر الجديدة", "جيت مول - الشروق", "ميدان هليوبوليس - مصر الجديدة", 
            "ش أحمد عرابى - المهندسين", "مستشفى السلام - التجمع الخامس", 
            "ش الخمسين - زهراء المعادى", "ستريب مول - مدينتى", "سانت تريزا - شبرا",
            "كرافت زون - مدينتى", "صيدلية الاتحاد - المنيا", "سيتى ستار مول - 6 أكتوبر", 
            "الحصرى - 6 أكتوبر"
        ]

        self.template_columns = [
            'Date', 'Time', 'Site', 'Transact. type', 'Article', 'Article Number', 
            'Trnsctn: Sales SUn', 'Gross Sales EGP', 'Sls Discount EGP', 'Net Sales EGP', 
            'POS no.', 'Trnsctn number', 'Manufacturer Name', 'offers Company', 
            'Activation date', 'End date'
        ]

        self.create_widgets()

    def create_widgets(self):
        title_text = f"Lotus Pro Claims (Mix & Match Engine) - v{APP_VERSION}"
        title_label = ctk.CTkLabel(self.root, text=title_text, font=ctk.CTkFont(size=24, weight="bold"))
        title_label.pack(pady=(10, 10))

        frame_file = ctk.CTkFrame(self.root)
        frame_file.pack(fill="x", padx=20, pady=5)
        
        ctk.CTkLabel(frame_file, text="1. Select Sales File", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(5, 0))
        
        btn_browse = ctk.CTkButton(frame_file, text="Browse File", command=self.browse_file)
        btn_browse.pack(side="left", padx=10, pady=10)
        
        self.lbl_file_path = ctk.CTkLabel(frame_file, textvariable=self.filepath, text_color="gray", wraplength=400)
        self.lbl_file_path.pack(side="left", padx=10, pady=10)
        
        btn_template = ctk.CTkButton(frame_file, text="⬇️ Download Template", fg_color="#2980B9", hover_color="#1F618D", command=self.download_template)
        btn_template.pack(side="right", padx=10, pady=10)

        frame_settings = ctk.CTkFrame(self.root)
        frame_settings.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(frame_settings, text="2. Filters & Target Discount", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(5, 5))

        ctk.CTkLabel(frame_settings, text="Branch / Site:").grid(row=1, column=0, sticky="e", padx=10, pady=5)
        self.combo_branch = ctk.CTkOptionMenu(frame_settings, values=self.branches, width=250)
        self.combo_branch.grid(row=1, column=1, sticky="w", padx=10, pady=5)

        ctk.CTkLabel(frame_settings, text="Target Redeem Disc (EGP):").grid(row=2, column=0, sticky="e", padx=10, pady=5)
        self.entry_target = ctk.CTkEntry(frame_settings, textvariable=self.target_discount_var, width=250, placeholder_text="e.g. 75000 (Empty = No Limit)")
        self.entry_target.grid(row=2, column=1, sticky="w", padx=10, pady=5)

        self.chk_date = ctk.CTkCheckBox(frame_settings, text="Filter by Date Range", variable=self.filter_by_date_var, command=self.toggle_dates)
        self.chk_date.grid(row=3, column=0, sticky="e", padx=10, pady=5)

        date_frame = ctk.CTkFrame(frame_settings, fg_color="transparent")
        date_frame.grid(row=3, column=1, sticky="w", padx=10)
        self.cal_start = DateEntry(date_frame, width=10, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', state="disabled")
        self.cal_start.pack(side="left", padx=5)
        ctk.CTkLabel(date_frame, text="to", text_color="gray").pack(side="left")
        self.cal_end = DateEntry(date_frame, width=10, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', state="disabled")
        self.cal_end.pack(side="left", padx=5)

        frame_mapping = ctk.CTkFrame(self.root)
        frame_mapping.pack(fill="both", expand=True, padx=20, pady=5)
        
        header_map = ctk.CTkFrame(frame_mapping, fg_color="transparent")
        header_map.pack(fill="x", padx=10, pady=(5, 5))
        ctk.CTkLabel(header_map, text="3. Map Offers to Companies (Mix & Match)", font=ctk.CTkFont(weight="bold")).pack(side="left")
        
        self.entry_new_offer = ctk.CTkEntry(header_map, width=150, placeholder_text="e.g. 3+1 or 20%")
        self.entry_new_offer.pack(side="left", padx=(30, 5))
        btn_add_offer = ctk.CTkButton(header_map, text="Add New Offer", width=120, fg_color="#F39C12", hover_color="#D35400", command=self.add_custom_offer)
        btn_add_offer.pack(side="left")

        ctrl_map = ctk.CTkFrame(frame_mapping, fg_color="transparent")
        ctrl_map.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(ctrl_map, text="Set ALL Companies to:").pack(side="left")
        self.combo_set_all = ctk.CTkOptionMenu(ctrl_map, values=self.available_offers, command=self.set_all_companies)
        self.combo_set_all.pack(side="left", padx=10)

        self.scroll_mapping = ctk.CTkScrollableFrame(frame_mapping, width=700, height=250)
        self.scroll_mapping.pack(padx=10, pady=5, fill="both", expand=True)
        ctk.CTkLabel(self.scroll_mapping, text="Upload an Excel/CSV file to load companies here.", text_color="gray").pack(pady=20)

        frame_action = ctk.CTkFrame(self.root, fg_color="transparent")
        frame_action.pack(pady=10)
        btn_process = ctk.CTkButton(frame_action, text="🚀 Process Mix & Match", font=ctk.CTkFont(weight="bold"), height=45, command=self.process_data)
        btn_process.pack(side="left", padx=10)
        btn_delete = ctk.CTkButton(frame_action, text="🗑️ Clear Data", fg_color="#E74C3C", hover_color="#C0392B", font=ctk.CTkFont(weight="bold"), height=45, command=self.clear_data)
        btn_delete.pack(side="left", padx=10)

    def download_template(self):
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="Lotus_Sales_Template.xlsx",
            title="Save Template File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if save_path:
            try:
                df_template = pd.DataFrame(columns=self.template_columns)
                df_template.to_excel(save_path, index=False, engine='openpyxl')
                messagebox.showinfo("Success", "Template downloaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template:\n{e}")

    def add_custom_offer(self):
        new_offer = self.entry_new_offer.get().strip().upper()
        if new_offer and new_offer not in self.available_offers:
            self.available_offers.append(new_offer)
            self.combo_set_all.configure(values=self.available_offers)
            for combo in self.company_checkbox_vars.values():
                combo.configure(values=self.available_offers)
            messagebox.showinfo("Added", f"Offer '{new_offer}' added successfully!")
            self.entry_new_offer.delete(0, 'end')

    def set_all_companies(self, selected_offer):
        for comp, var in self.company_mappings.items():
            var.set(selected_offer)

    def toggle_dates(self):
        if self.filter_by_date_var.get():
            self.cal_start.config(state="normal")
            self.cal_end.config(state="normal")
        else:
            self.cal_start.config(state="disabled")
            self.cal_end.config(state="disabled")

    def build_mapping_ui(self, companies):
        for widget in self.scroll_mapping.winfo_children():
            widget.destroy()
        self.company_mappings.clear()
        self.company_checkbox_vars.clear()

        if not companies:
            ctk.CTkLabel(self.scroll_mapping, text="No companies found in this file.").pack()
            return

        for comp in companies:
            comp_clean = str(comp).strip()
            if comp_clean == "" or comp_clean.lower() == 'nan': continue
            if comp_clean in self.company_mappings: continue
            
            row_frame = ctk.CTkFrame(self.scroll_mapping, fg_color="transparent")
            row_frame.pack(fill="x", pady=2)
            
            lbl = ctk.CTkLabel(row_frame, text=comp_clean, width=300, anchor="w", font=ctk.CTkFont(weight="bold"))
            lbl.pack(side="left", padx=5)
            
            offer_var = ctk.StringVar(value="1+50%")
            combo = ctk.CTkOptionMenu(row_frame, variable=offer_var, values=self.available_offers, width=200)
            combo.pack(side="left", padx=10)
            
            self.company_mappings[comp_clean] = offer_var
            self.company_checkbox_vars[comp_clean] = combo

    def browse_file(self):
        filename = filedialog.askopenfilename(title="Select Sales File", filetypes=(("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")))
        if filename:
            self.filepath.set(filename)
            try:
                if filename.endswith(('.xlsx', '.xls')):
                    df = pd.read_excel(filename)
                else:
                    try:
                        df = pd.read_csv(filename, encoding='utf-8-sig')
                    except UnicodeDecodeError:
                        df = pd.read_csv(filename, encoding='cp1256')
                
                df.columns = df.columns.str.strip()
                if 'Manufacturer Name' in df.columns:
                    companies = sorted(df['Manufacturer Name'].dropna().astype(str).unique().tolist())
                    self.build_mapping_ui(companies)
                else:
                    messagebox.showwarning("Warning", "Column 'Manufacturer Name' not found in file!")
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file:\n{e}")

    def clear_data(self):
        self.filepath.set("")
        self.combo_branch.set("All")
        self.target_discount_var.set("")
        self.filter_by_date_var.set(False)
        self.toggle_dates()
        for widget in self.scroll_mapping.winfo_children():
            widget.destroy()
        self.company_mappings.clear()
        self.company_checkbox_vars.clear()
        ctk.CTkLabel(self.scroll_mapping, text="Upload an Excel/CSV file to load companies here.", text_color="gray").pack(pady=20)

    def process_data(self):
        input_file = self.filepath.get()
        if not input_file:
            messagebox.showwarning("Warning", "Please select a file first!")
            return

        try:
            target_str = self.target_discount_var.get().strip()
            target_discount = float(target_str) if target_str else float('inf')
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number for Target Discount!")
            return

        try:
            if input_file.endswith(('.xlsx', '.xls')):
                df_original = pd.read_excel(input_file)
            else:
                try:
                    df_original = pd.read_csv(input_file, encoding='utf-8-sig')
                except UnicodeDecodeError:
                    df_original = pd.read_csv(input_file, encoding='cp1256')

            df_original.columns = df_original.columns.str.strip()
            df_original['Date'] = pd.to_datetime(df_original['Date']).dt.date
            
            if 'Time' in df_original.columns:
                df_original['Time'] = pd.to_datetime(df_original['Time'].astype(str), errors='coerce').dt.strftime('%I:%M:%S %p').fillna(df_original['Time'])
            
            df_original['Site'] = df_original['Site'].fillna('Unknown')
            if 'Transact. type' in df_original.columns:
                df_original['Transact. type'] = df_original['Transact. type'].fillna('Unknown')

            dedupe_cols = ['Date', 'Site', 'Trnsctn number', 'POS no.']
            if 'Article Number' in df_original.columns:
                dedupe_cols.append('Article Number')
            elif 'Article' in df_original.columns:
                dedupe_cols.append('Article')
            df = df_original.drop_duplicates(subset=dedupe_cols, keep='first').copy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the data:\n{e}")
            return

        target_branch = self.combo_branch.get()
        use_date_filter = self.filter_by_date_var.get()
        
        if use_date_filter:
            try:
                start_date = self.cal_start.get_date()
                end_date = self.cal_end.get_date()
            except Exception:
                messagebox.showerror("Error", "Please select valid dates from the calendar!")
                return
        else:
            start_date = None
            end_date = None

        company_mappings = {comp: var.get() for comp, var in self.company_mappings.items()}
        processed_data, true_unprocessed, accumulated_discount = process_mix_match(
            df_original,
            company_mappings,
            target_branch=target_branch,
            use_date_filter=use_date_filter,
            start_date=start_date,
            end_date=end_date,
            target_discount=target_discount,
        )

        total_processed = sum(len(lst) for lst in processed_data.values())
        if total_processed == 0 and not true_unprocessed:
            messagebox.showinfo("Info", "No data generated. Check your mappings.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Lotus_MixMatch_v{APP_VERSION}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            title="Save Processed File",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            excel_bytes = export_to_excel(processed_data, true_unprocessed)
            with open(save_path, 'wb') as out:
                out.write(excel_bytes)

            target_display = target_str if target_str else "No Limit"
            msg = (f"🎉 Mix & Match Exported Successfully!\n\n"
                   f"Version: {APP_VERSION}\n"
                   f"Target Discount: {target_display}\n"
                   f"Achieved Total Discount: {accumulated_discount:,.2f} EGP\n\n"
                   f"Unprocessed Items (Single Leftovers & Skips): {len(true_unprocessed)}")
            messagebox.showinfo("Success", msg)

if __name__ == "__main__":
    app_window = ctk.CTk()
    app = LotusOfferGenerator(app_window)
    app_window.mainloop()