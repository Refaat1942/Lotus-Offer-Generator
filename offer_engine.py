"""Lotus Offer Generator - Mix & Match Engine (Core Logic)
Mirrors Offer.py v1.4 process_data logic exactly."""
import io
import os
import random
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font

APP_VERSION = "1.4"

TEMPLATE_COLUMNS = [
    'Date', 'Time', 'Site', 'Transact. type', 'Article', 'Article Number',
    'Trnsctn: Sales SUn', 'Gross Sales EGP', 'Sls Discount EGP', 'Net Sales EGP',
    'POS no.', 'Trnsctn number', 'Manufacturer Name', 'offers Company',
    'Activation date', 'End date'
]

BRANCHES = [
    "All", "حي السفارات - مدينة نصر", "مول المعز - الشيخ زايد",
    "السوق الشرقي - الرحاب", "الرؤى مول - العاشر من رمضان",
    "تريومف - مصر الجديدة", "جيت مول - الشروق", "ميدان هليوبوليس - مصر الجديدة",
    "ش أحمد عرابى - المهندسين", "مستشفى السلام - التجمع الخامس",
    "ش الخمسين - زهراء المعادى", "ستريب مول - مدينتى", "سانت تريزا - شبرا",
    "كرافت زون - مدينتى", "صيدلية الاتحاد - المنيا", "سيتى ستار مول - 6 أكتوبر",
    "الحصرى - 6 أكتوبر"
]

DEFAULT_OFFERS = ["Skip", "1+1", "1+50%", "2+1", "10%", "20%"]


def parse_offer_string(offer_str):
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
    except Exception:
        return None


def normalize_manufacturer_name(name):
    """Normalize manufacturer keys so UI mappings match Excel groupby values."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ''
    s = str(name).strip()
    if not s or s.lower() == 'nan':
        return ''
    if s.endswith('.0'):
        try:
            s = str(int(float(s)))
        except ValueError:
            pass
    return s


def build_mapping_lookup(company_mappings):
    """Build lookup with both exact and normalized manufacturer keys."""
    lookup = {}
    for key, offer in (company_mappings or {}).items():
        exact = str(key).strip()
        if exact:
            lookup[exact] = offer
        norm = normalize_manufacturer_name(key)
        if norm:
            lookup.setdefault(norm, offer)
    return lookup


def resolve_offer(lookup, manufacturer_value, default="Skip"):
    key = str(manufacturer_value).strip()
    if key in lookup:
        return lookup[key]
    return lookup.get(normalize_manufacturer_name(manufacturer_value), default)


def _read_excel_simple(file_bytes):
    """Same as Offer.py: pd.read_excel then strip column names."""
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = df.columns.str.strip()
    return df


def _find_header_row(raw_df):
    markers = {'Manufacturer Name', 'Date', 'Site', 'Trnsctn number', 'Gross Sales EGP'}
    for i in range(min(25, len(raw_df))):
        row_vals = {str(v).strip() for v in raw_df.iloc[i].values if pd.notna(v) and str(v).strip()}
        if 'Manufacturer Name' in row_vals or ('Date' in row_vals and 'Site' in row_vals):
            return i
        if len(markers.intersection(row_vals)) >= 2:
            return i
    return 0


def _read_excel_with_header_detect(file_bytes):
    """Fallback when simple read does not find expected columns."""
    buf = io.BytesIO(file_bytes)
    raw = pd.read_excel(buf, header=None, engine='openpyxl')
    header_idx = _find_header_row(raw)
    buf.seek(0)
    df = pd.read_excel(buf, header=header_idx, engine='openpyxl')
    df.columns = df.columns.str.strip()
    return df.dropna(how='all')


def read_sales_file(file_bytes, filename):
    """Read sales file — primary path matches Offer.py v1.4."""
    ext = os.path.splitext((filename or '').lower())[1]

    if ext in ('.xlsx', '.xls', '.xlsm', '.xlsb'):
        try:
            df = _read_excel_simple(file_bytes)
        except Exception:
            df = _read_excel_with_header_detect(file_bytes)
        if 'Manufacturer Name' not in df.columns and len(df.columns) > 0:
            df = _read_excel_with_header_detect(file_bytes)
    else:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding='cp1256')
        df.columns = df.columns.str.strip()

    return df


def prepare_dataframe(df_original):
    """Exact preprocessing from Offer.py v1.4 process_data (lines 327-337)."""
    df_original = df_original.copy()
    df_original.columns = df_original.columns.str.strip()
    df_original['Date'] = pd.to_datetime(df_original['Date']).dt.date

    if 'Time' in df_original.columns:
        df_original['Time'] = pd.to_datetime(
            df_original['Time'].astype(str), errors='coerce'
        ).dt.strftime('%I:%M:%S %p').fillna(df_original['Time'])

    df_original['Site'] = df_original['Site'].fillna('Unknown')

    if 'Transact. type' not in df_original.columns:
        df_original['Transact. type'] = 'Unknown'
    else:
        df_original['Transact. type'] = df_original['Transact. type'].fillna('Unknown')

    df = df_original.drop_duplicates(
        subset=['Date', 'Site', 'Trnsctn number', 'POS no.', 'Article'], keep='first'
    ).copy()
    return df_original, df


def get_companies_from_df(df):
    """Same company list logic as Offer.py browse_file + build_mapping_ui."""
    if 'Manufacturer Name' not in df.columns:
        return []
    companies = sorted(df['Manufacturer Name'].dropna().astype(str).unique().tolist())
    result = []
    seen = set()
    for comp in companies:
        comp_clean = str(comp).strip()
        if comp_clean == '' or comp_clean.lower() == 'nan':
            continue
        if comp_clean in seen:
            continue
        seen.add(comp_clean)
        result.append(comp_clean)
    return result


def create_template_bytes():
    df_template = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    buffer = io.BytesIO()
    df_template.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)
    return buffer.getvalue()


def process_mix_match(
    df_original,
    company_mappings,
    target_branch="All",
    use_date_filter=False,
    start_date=None,
    end_date=None,
    target_discount=float('inf'),
):
    df_original, df = prepare_dataframe(df_original)
    mapping_lookup = build_mapping_lookup(company_mappings)

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
    used_signatures = set()
    valid_pos_list = ['101', '102', '103', '104']

    def get_safe_pos_and_time(b_date, b_site):
        while True:
            curr_pos = random.choice(valid_pos_list)
            hour = random.randint(9, 23)
            minute = random.randint(0, 59)
            second = random.randint(0, 59)
            time_obj = datetime.strptime(f"{hour}:{minute}:{second}", "%H:%M:%S")
            curr_time = time_obj.strftime('%I:%M:%S %p')
            if (b_date, b_site, curr_pos, curr_time) not in used_signatures:
                used_signatures.add((b_date, b_site, curr_pos, curr_time))
                return curr_pos, curr_time

    def add_bundle(offer_name, bundle, p_count, d_val, pct_flag, is_native=False):
        nonlocal accumulated_discount
        bundle.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
        p_items = bundle[:p_count]
        d_items = bundle[p_count:]

        pot_disc = 0.0
        for item in d_items:
            price = float(item['Gross Sales EGP'])
            pot_disc += round(price * d_val, 2) if pct_flag else price

        apply_disc = (accumulated_discount + pot_disc <= target_discount) or (target_discount == float('inf'))
        if apply_disc:
            accumulated_discount += pot_disc

        if is_native:
            n_date = bundle[0]['Date']
            n_site = bundle[0]['Site']
            n_pos = str(bundle[0].get('POS no.', '101')).strip()
            n_time = str(bundle[0].get('Time', '12:00:00 PM')).strip()
            used_signatures.add((n_date, n_site, n_pos, n_time))

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
            base_item = bundle[0]
            base_date = base_item['Date']
            base_site = base_item['Site']
            safe_pos, safe_time = get_safe_pos_and_time(base_date, base_site)
            new_t_num = str(get_new_t_num(base_date, base_site))

            for is_discounted, items_subset in [(False, p_items), (True, d_items)]:
                for item in items_subset:
                    r = item.copy()
                    r['Date'] = base_date
                    r['Site'] = base_site
                    r['POS no.'] = safe_pos
                    r['Time'] = safe_time
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

    grouped = df.groupby('Manufacturer Name', dropna=False)

    for comp_val, group in grouped:
        comp_val_str = str(comp_val).strip()
        offer_str = resolve_offer(mapping_lookup, comp_val, "Skip")

        if offer_str == "Skip":
            true_unprocessed.extend(group.to_dict('records'))
            continue

        parsed_offer = parse_offer_string(offer_str)
        if not parsed_offer:
            true_unprocessed.extend(group.to_dict('records'))
            continue

        chunk_size, paid_count, discount_val, is_percentage = parsed_offer

        for sale_type, type_group in group.groupby('Transact. type', dropna=False):
            company_type_leftovers = []

            for (d_val, s_val, t_val), trans_group in type_group.groupby(['Date', 'Site', 'Trnsctn number']):
                trans_items = []
                for _, row in trans_group.iterrows():
                    qty = float(str(row.get('Trnsctn: Sales SUn', 0)).replace(',', '').strip() or 0)
                    if qty <= 0:
                        true_unprocessed.append(row.to_dict())
                        continue
                    unit_gross = round(
                        float(str(row.get('Gross Sales EGP', 0)).replace(',', '').strip() or 0) / qty, 2
                    )
                    for _ in range(int(qty)):
                        item = row.to_dict()
                        item['Trnsctn: Sales SUn'] = 1.0
                        item['Gross Sales EGP'] = unit_gross
                        trans_items.append(item)

                trans_items.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
                K = len(trans_items) // chunk_size
                company_type_leftovers.extend(trans_items[K * chunk_size:])

                for b_idx in range(K):
                    bundle = trans_items[b_idx * chunk_size: (b_idx + 1) * chunk_size]
                    is_native = (b_idx == 0)
                    add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage, is_native=is_native)

            if company_type_leftovers:
                company_type_leftovers.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
                K_left = len(company_type_leftovers) // chunk_size
                final_orphans = company_type_leftovers[K_left * chunk_size:]

                for b_idx in range(K_left):
                    bundle = company_type_leftovers[b_idx * chunk_size: (b_idx + 1) * chunk_size]
                    add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage, is_native=False)

                while len(final_orphans) >= 2:
                    bundle = final_orphans[:2]
                    final_orphans = final_orphans[2:]
                    add_bundle("1+50%", bundle, 1, 0.5, True, is_native=False)

                for item in final_orphans:
                    r = item.copy()
                    r['Sls Discount EGP'] = 0.00
                    r['Net Sales EGP'] = float(r['Gross Sales EGP'])
                    true_unprocessed.append(r)

    return processed_data, true_unprocessed, accumulated_discount


def export_to_excel(processed_data, true_unprocessed):
    buffer = io.BytesIO()
    bold_blue_font = Font(bold=True, color="0000FF")

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for offer_name, rows_list in processed_data.items():
            if rows_list:
                df_proc = pd.DataFrame(rows_list)
                final_cols = [col for col in TEMPLATE_COLUMNS if col in df_proc.columns]
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
            final_cols_unp = [col for col in TEMPLATE_COLUMNS if col in df_unp.columns]
            df_unp[final_cols_unp].to_excel(writer, index=False, sheet_name='Unprocessed')

    buffer.seek(0)
    return buffer.getvalue()
