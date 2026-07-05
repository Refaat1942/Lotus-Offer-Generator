import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font

# إعدادات المظهر
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class LotusOfferGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Lotus Pro Claims - Mix & Match Engine")
        self.root.geometry("850x850")
        
        self.filepath = ctk.StringVar()
        self.filter_by_date_var = ctk.BooleanVar(value=False)
        self.target_discount_var = ctk.StringVar()
        
        self.available_offers = ["Skip", "1+1", "1+50%", "2+1", "10%", "20%"]
        
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
        title_label = ctk.CTkLabel(self.root, text="Lotus Pro Claims (Mix & Match Engine)", font=ctk.CTkFont(size=24, weight="bold"))
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
            
            offer_var = ctk.StringVar(value="Skip")
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

    def parse_offer_string(self, offer_str):
        offer_str = offer_str.replace(" ", "")
        try:
            if '+' in offer_str:
                parts = offer_str.split('+')
                if '%' in parts[1]:
                    x = int(parts[0])
                    pct = float(parts[1].replace('%', '')) / 100.0
                    return (x + 1, x, pct, True)
                else:
                    x = int(parts[0])
                    y = int(parts[1])
                    return (x + y, x, 1.0, False)
            elif '%' in offer_str:
                pct = float(offer_str.replace('%', '')) / 100.0
                return (2, 1, min(pct * 2, 1.0), True)
            else:
                return None
        except:
            return None

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
            df = df_original.drop_duplicates(subset=['Date', 'Site', 'Trnsctn number', 'POS no.', 'Article'], keep='first').copy()
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
            
        # استخراج الأرقام الساقطة من الشيت الأصلي
        site_date_trans_info = {}
        for (d, s), group in df_original.groupby(['Date', 'Site']):
            nums = pd.to_numeric(group['Trnsctn number'], errors='coerce').dropna().astype(int)
            if not nums.empty:
                t_min, t_max = nums.min(), nums.max()
                if t_max - t_min < 100000:
                    t_set = set(nums)
                    missing = sorted(list(set(range(t_min, t_max + 1)) - t_set))
                else:
                    missing = []
                site_date_trans_info[(d, s)] = {'missing': missing, 'max': t_max}
            else:
                site_date_trans_info[(d, s)] = {'missing': [], 'max': 1000}

        def get_new_t_num(date_v, site_v):
            info = site_date_trans_info.get((date_v, site_v), {'missing': [], 'max': 1000})
            if info['missing']:
                new_val = info['missing'].pop(0)
            else:
                info['max'] += 1
                new_val = info['max']
            site_date_trans_info[(date_v, site_v)] = info
            return str(new_val)

        if target_branch != "All":
            df = df[df['Site'].str.strip() == target_branch.strip()]
        if use_date_filter:
            df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        processed_data = {} 
        true_unprocessed = []
        accumulated_discount = 0.0

        # ---- محرك توحيد وتأليف العروض الصارم ----
        def add_bundle(offer_name, bundle, p_count, d_val, pct_flag, is_native=False):
            nonlocal accumulated_discount
            
            # ترتيب الأصناف من الأغلى للأرخص إجبارياً
            bundle.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
            p_items = bundle[:p_count] # الأصناف الأغلى (هتندفع بالكامل)
            d_items = bundle[p_count:] # الأصناف الأرخص (هتاخد الخصم)
            
            pot_disc = 0.0
            for item in d_items:
                price = float(item['Gross Sales EGP'])
                pot_disc += round(price * d_val, 2) if pct_flag else price
                
            apply_disc = (accumulated_discount + pot_disc <= target_discount) or (target_discount == float('inf'))
            if apply_disc:
                accumulated_discount += pot_disc
                
            if is_native:
                # لو القطع دي جاية جاهزة من نفس الفرع والتاريخ، متلعبش في بياناتهم بس اخصم صح
                for is_discounted, items_subset in [(False, p_items), (True, d_items)]:
                    for item in items_subset:
                        r = item.copy()
                        r['Is_New_Trans'] = False
                        price = float(r['Gross Sales EGP'])
                        if is_discounted and apply_disc:
                            item_disc = round(price * d_val, 2) if pct_flag else price
                            r['Sls Discount EGP'] = -item_disc
                            r['Net Sales EGP'] = price - item_disc
                        else:
                            r['Sls Discount EGP'] = 0.00
                            r['Net Sales EGP'] = price
                        processed_data.setdefault(offer_name, []).append(r)
            else:
                # هنا السر: التوحيد الإجباري!
                # بناخد بيانات أغلى قطعة (الفرع، التاريخ، رقم المكنة، الوقت) ونمشيها على كل القطع
                base_item = bundle[0]
                base_date = base_item['Date']
                base_site = base_item['Site']
                base_pos = base_item.get('POS no.', 'Unknown')
                base_time = base_item.get('Time', '')
                new_t_num = str(get_new_t_num(base_date, base_site))
                
                for is_discounted, items_subset in [(False, p_items), (True, d_items)]:
                    for item in items_subset:
                        r = item.copy()
                        
                        # توحيد كامل لكل البيانات عشان تظهر في الإكسيل كأنها فاتورة واحدة فعلاً 
                        r['Date'] = base_date
                        r['Site'] = base_site
                        r['POS no.'] = base_pos
                        r['Time'] = base_time
                        r['Trnsctn number'] = new_t_num
                        r['Is_New_Trans'] = True
                        
                        price = float(r['Gross Sales EGP'])
                        if is_discounted and apply_disc:
                            item_disc = round(price * d_val, 2) if pct_flag else price
                            r['Sls Discount EGP'] = -item_disc
                            r['Net Sales EGP'] = price - item_disc
                        else:
                            r['Sls Discount EGP'] = 0.00
                            r['Net Sales EGP'] = price
                        
                        processed_data.setdefault(offer_name, []).append(r)
        # ------------------------------------------------------------------

        grouped = df.groupby('Manufacturer Name', dropna=False)

        for comp_val, group in grouped:
            comp_val_str = str(comp_val).strip()
            offer_str = self.company_mappings.get(comp_val_str, ctk.StringVar(value="Skip")).get()

            if offer_str == "Skip":
                true_unprocessed.extend(group.to_dict('records'))
                continue

            parsed_offer = self.parse_offer_string(offer_str)
            if not parsed_offer:
                true_unprocessed.extend(group.to_dict('records'))
                continue
                
            chunk_size, paid_count, discount_val, is_percentage = parsed_offer
            company_leftovers = []

            # 1. المرحلة الأولى: استغلال الأصناف اللي جاية جاهزة في نفس الفاتورة
            for (d_val, s_val, t_val), trans_group in group.groupby(['Date', 'Site', 'Trnsctn number']):
                trans_items = []
                for _, row in trans_group.iterrows():
                    qty = float(str(row.get('Trnsctn: Sales SUn', 0)).replace(',', '').strip() or 0)
                    if qty <= 0:
                        true_unprocessed.append(row.to_dict())
                        continue
                    unit_gross = round(float(str(row.get('Gross Sales EGP', 0)).replace(',', '').strip() or 0) / qty, 2)
                    for _ in range(int(qty)):
                        item = row.to_dict()
                        item['Trnsctn: Sales SUn'] = 1.0
                        item['Gross Sales EGP'] = unit_gross
                        trans_items.append(item)
                        
                trans_items.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
                
                K = len(trans_items) // chunk_size
                company_leftovers.extend(trans_items[K * chunk_size:])
                
                for b_idx in range(K):
                    bundle = trans_items[b_idx * chunk_size : (b_idx + 1) * chunk_size]
                    is_native = (b_idx == 0) # أول حزمة تفضل برقم فاتورتها، التانية تتفصل بفاتورة جديدة في نفس الفرع
                    add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage, is_native=is_native)
                    
            # 2. المرحلة التانية: تجميع كل البواقي (من فروع وتواريخ مختلفة) وتأليف عروض موحدة
            if company_leftovers:
                company_leftovers.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
                K_left = len(company_leftovers) // chunk_size
                final_orphans = company_leftovers[K_left * chunk_size:]
                
                for b_idx in range(K_left):
                    bundle = company_leftovers[b_idx * chunk_size : (b_idx + 1) * chunk_size]
                    # نبعت العرض يتوحد ويبقى فاتورة شرعية
                    add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage, is_native=False)
                    
                # 3. مرحلة الطوارئ: لو اتبقى قطعتين في النهاية نعمل بيهم فاتورة (1+50%) موحدة
                while len(final_orphans) >= 2:
                    bundle = final_orphans[:2]
                    final_orphans = final_orphans[2:]
                    add_bundle("1+50%", bundle, 1, 0.5, True, is_native=False)
                    
                # 4. لو اتبقت قطعة واحدة يتيمة مستحيل تاخد خصم (تترمي Unprocessed زي ما هي)
                for item in final_orphans:
                    r = item.copy()
                    r['Sls Discount EGP'] = 0.00
                    r['Net Sales EGP'] = float(r['Gross Sales EGP'])
                    true_unprocessed.append(r)

        total_processed = sum(len(lst) for lst in processed_data.values())
        if total_processed == 0 and not true_unprocessed:
            messagebox.showinfo("Info", "No data generated. Check your mappings.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Lotus_MixMatch_{datetime.now().strftime('%Y%m%d')}.xlsx",
            title="Save Processed File",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            bold_blue_font = Font(bold=True, color="0000FF")
            
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for offer_name, rows_list in processed_data.items():
                    if rows_list:
                        df_proc = pd.DataFrame(rows_list)
                        final_cols = [col for col in self.template_columns if col in df_proc.columns]
                        safe_sheet_name = f'Offer_{offer_name}'.replace('%', 'pct') 
                        df_proc[final_cols].to_excel(writer, index=False, sheet_name=safe_sheet_name)
                        
                        worksheet = writer.sheets[safe_sheet_name]
                        if 'Trnsctn number' in final_cols:
                            t_num_idx = final_cols.index('Trnsctn number') + 1
                            is_new_list = df_proc.get('Is_New_Trans', [False] * len(df_proc)).tolist()
                            
                            for row_idx, is_new in enumerate(is_new_list, start=2): 
                                if is_new:
                                    worksheet.cell(row=row_idx, column=t_num_idx).font = bold_blue_font

                if true_unprocessed:
                    df_unp = pd.DataFrame(true_unprocessed)
                    final_cols_unp = [col for col in self.template_columns if col in df_unp.columns]
                    df_unp[final_cols_unp].to_excel(writer, index=False, sheet_name='Unprocessed')

            target_display = target_str if target_str else "No Limit"
            msg = (f"🎉 Mix & Match Exported Successfully!\n\n"
                   f"Target Discount: {target_display}\n"
                   f"Achieved Total Discount: {accumulated_discount:,.2f} EGP\n\n"
                   f"Unprocessed Items (Single Leftovers & Skips): {len(true_unprocessed)}")
            messagebox.showinfo("Success", msg)

if __name__ == "__main__":
    app_window = ctk.CTk()
    app = LotusOfferGenerator(app_window)
    app_window.mainloop()