"""Lotus Offer Generator - Mix & Match Engine (Core Logic)"""
import io
import os
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

DEFAULT_OFFERS = ["Skip", "1+50%", "1+1", "2+1", "10%", "20%"]


def parse_offer_string(offer_str):
    """
    Parse company offer from the process page.

    Returns:
      ('pair', chunk_size, paid_count, discount_val, is_percentage)
          Mix & match: 1+50%, 1+1, 2+1, 1+40% → two lines per bundle
      ('direct', discount_pct)
          Straight discount: 10%, 20% → discount on the same line
      None if invalid
    """
    offer_str = offer_str.replace(" ", "")
    try:
        if '+' in offer_str:
            parts = offer_str.split('+')
            if '%' in parts[1]:
                x = int(parts[0])
                pct = float(parts[1].replace('%', '')) / 100.0
                return ('pair', x + 1, x, pct, True)
            x = int(parts[0])
            y = int(parts[1])
            return ('pair', x + y, x, 1.0, False)
        if '%' in offer_str:
            pct = float(offer_str.replace('%', '')) / 100.0
            return ('direct', pct)
        return None
    except Exception:
        return None


def is_pair_offer(parsed):
    return bool(parsed and parsed[0] == 'pair')


def is_direct_offer(parsed):
    return bool(parsed and parsed[0] == 'direct')


def unpack_pair_offer(parsed):
    if not is_pair_offer(parsed):
        return None
    return parsed[1], parsed[2], parsed[3], parsed[4]


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
        subset=_dedupe_columns(df_original), keep='first'
    ).copy()
    return df_original, df


def _dedupe_columns(df):
    """Dedupe line items without collapsing different products on the same receipt."""
    cols = ['Date', 'Site', 'Trnsctn number', 'POS no.']
    if 'Article Number' in df.columns:
        cols.append('Article Number')
    elif 'Article' in df.columns:
        cols.append('Article')
    return cols


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


def _item_fingerprint(item):
    """Unique key for a single sale unit (prevents duplicate export rows)."""
    gross = item.get('Gross Sales EGP', 0)
    try:
        gross = round(float(str(gross).replace(',', '').strip() or 0), 2)
    except (TypeError, ValueError):
        gross = 0.0
    article_key = item.get('Article Number', item.get('Article', ''))
    return (
        str(item.get('Date', '')),
        str(item.get('Site', '')).strip(),
        str(item.get('POS no.', '')).strip(),
        normalize_manufacturer_name(item.get('Manufacturer Name')),
        str(article_key).strip(),
        gross,
    )


def _collect_processed_fingerprints(processed_data):
    fps = set()
    for rows in processed_data.values():
        for row in rows:
            fps.add(_item_fingerprint(row))
    return fps


def _filter_unprocessed_export(true_unprocessed, processed_data):
    """Drop rows already represented in processed sheets (keeps quantities correct)."""
    seen = _collect_processed_fingerprints(processed_data)
    filtered = []
    for row in true_unprocessed:
        fp = _item_fingerprint(row)
        if fp in seen:
            continue
        filtered.append(row)
        seen.add(fp)
    return filtered


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

    def bundle_shares_same_register(bundle):
        """True when every item is on the same date, branch, and POS register."""
        if not bundle:
            return False
        lead = bundle[0]
        reg_date = lead['Date']
        reg_site = lead['Site']
        reg_pos = str(lead.get('POS no.', '101')).strip()
        return all(
            item['Date'] == reg_date
            and item['Site'] == reg_site
            and str(item.get('POS no.', '101')).strip() == reg_pos
            for item in bundle
        )

    def unified_receipt_fields(bundle):
        """Align paired lines to one receipt: date, POS, time, transaction, transact type."""
        lead = bundle[0]
        reg_date = lead['Date']
        reg_site = lead['Site']
        reg_pos = str(lead.get('POS no.', '101')).strip()
        reg_time = str(lead.get('Time', '12:00:00 PM')).strip()
        reg_t_num = str(lead.get('Trnsctn number', '')).strip()
        reg_transact = lead.get('Transact. type', 'Unknown')

        if bundle_shares_same_register(bundle):
            return reg_date, reg_site, reg_pos, reg_time, reg_t_num, reg_transact, False

        new_t_num = str(get_new_t_num(reg_date, reg_site))
        return reg_date, reg_site, reg_pos, reg_time, new_t_num, reg_transact, True

    def write_bundle_rows(offer_name, bundle, p_items, d_items, apply_disc, d_val, pct_flag,
                          u_date=None, u_site=None, u_pos=None, u_time=None, u_t_num=None,
                          u_transact=None, is_new_trans=False, align_fields=False):
        for is_discounted, items_subset in [(False, p_items), (True, d_items)]:
            for item in items_subset:
                r = item.copy()
                if align_fields:
                    r['Date'] = u_date
                    r['Site'] = u_site
                    r['POS no.'] = u_pos
                    r['Time'] = u_time
                    r['Trnsctn number'] = u_t_num
                    if u_transact is not None:
                        r['Transact. type'] = u_transact
                r['Is_New_Trans'] = is_new_trans
                r['offers Company'] = offer_name
                price = float(r['Gross Sales EGP'])
                if is_discounted and apply_disc:
                    item_disc = round(price * d_val, 2) if pct_flag else price
                    r['Sls Discount EGP'] = -item_disc
                    r['Net Sales EGP'] = price - item_disc
                else:
                    r['Sls Discount EGP'] = 0.00
                    r['Net Sales EGP'] = price
                processed_data.setdefault(offer_name, []).append(r)

    def add_bundle(offer_name, bundle, p_count, d_val, pct_flag):
        """Mix & match bundle → exactly len(bundle) lines, same receipt on each line."""
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

        u_date, u_site, u_pos, u_time, u_t_num, u_transact, is_new_trans = unified_receipt_fields(bundle)
        write_bundle_rows(
            offer_name, bundle, p_items, d_items, apply_disc, d_val, pct_flag,
            u_date, u_site, u_pos, u_time, u_t_num, u_transact, is_new_trans, align_fields=True,
        )

    def add_pair_bundle(offer_str, bundle):
        """Pair exactly 2 items using the mapped mix & match offer."""
        parsed = parse_offer_string(offer_str)
        pair = unpack_pair_offer(parsed)
        if pair and len(bundle) >= 2:
            chunk_size, paid_count, discount_val, is_percentage = pair
            add_bundle(offer_str, bundle[:chunk_size], paid_count, discount_val, is_percentage)
        elif len(bundle) >= 2:
            add_bundle('1+50%', bundle[:2], 1, 0.5, True)

    def process_direct_discount(group, offer_str, discount_pct):
        """10% / 20% — one output line per source row, discount on same line."""
        nonlocal accumulated_discount
        for _, row in group.iterrows():
            r = row.to_dict()
            qty = float(str(r.get('Trnsctn: Sales SUn', 0)).replace(',', '').strip() or 0)
            if qty <= 0:
                true_unprocessed.append(r)
                continue
            price = float(str(r.get('Gross Sales EGP', 0)).replace(',', '').strip() or 0)
            pot_disc = round(price * discount_pct, 2)
            apply_disc = (accumulated_discount + pot_disc <= target_discount) or (target_discount == float('inf'))
            r['offers Company'] = offer_str
            r['Is_New_Trans'] = False
            if apply_disc:
                r['Sls Discount EGP'] = -pot_disc
                r['Net Sales EGP'] = round(price - pot_disc, 2)
                accumulated_discount += pot_disc
            else:
                r['Sls Discount EGP'] = 0.00
                r['Net Sales EGP'] = price
            processed_data.setdefault(offer_str, []).append(r)

    def bucket_items(items, key_fn):
        buckets = {}
        for item in items:
            buckets.setdefault(key_fn(item), []).append(item)
        return buckets

    def process_item_pool(items, offer_str, chunk_size, paid_count, discount_val, is_percentage):
        """Form offer bundles from a pool; return items that could not be paired."""
        if not items:
            return []

        items.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
        k = len(items) // chunk_size
        orphans = items[k * chunk_size:]
        for b_idx in range(k):
            bundle = items[b_idx * chunk_size: (b_idx + 1) * chunk_size]
            add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage)
        return orphans

    def process_leftover_pool(items, offer_str, chunk_size, paid_count, discount_val, is_percentage):
        """Pair leftovers: same receipt, same POS register, branch-wide, then pairs of 2."""
        if not items:
            return

        branch_leftovers = []

        # 1) Exact same receipt (date + branch + POS + transaction)
        for receipt_items in bucket_items(
            items,
            lambda it: (
                it['Date'], it['Site'],
                str(it.get('POS no.', '101')).strip(),
                str(it.get('Trnsctn number', '')).strip(),
            ),
        ).values():
            branch_leftovers.extend(
                process_item_pool(receipt_items, offer_str, chunk_size, paid_count, discount_val, is_percentage)
            )

        # 2) Same register/day (date + branch + POS) — unify to one receipt line pattern
        pos_leftovers = []
        for pos_items in bucket_items(
            branch_leftovers,
            lambda it: (it['Date'], it['Site'], str(it.get('POS no.', '101')).strip()),
        ).values():
            pos_leftovers.extend(
                process_item_pool(pos_items, offer_str, chunk_size, paid_count, discount_val, is_percentage)
            )

        pos_leftovers.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
        k_branch = len(pos_leftovers) // chunk_size
        final_orphans = pos_leftovers[k_branch * chunk_size:]

        for b_idx in range(k_branch):
            bundle = pos_leftovers[b_idx * chunk_size: (b_idx + 1) * chunk_size]
            add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage)

        while len(final_orphans) >= 2:
            bundle = final_orphans[:2]
            final_orphans = final_orphans[2:]
            add_pair_bundle(offer_str, bundle)

        # Single leftovers are omitted — exporting them duplicates source qty counts

    def sweep_unprocessed_receipt_pairs():
        """Last pass: pair any same-POS orphans still sitting in Unprocessed."""
        nonlocal true_unprocessed
        pending = []
        pair_groups = {}

        for item in true_unprocessed:
            offer_str = resolve_offer(mapping_lookup, item.get('Manufacturer Name'), 'Skip')
            if offer_str == 'Skip':
                pending.append(item)
                continue
            parsed = parse_offer_string(offer_str)
            if not is_pair_offer(parsed):
                pending.append(item)
                continue

            key = (
                normalize_manufacturer_name(item.get('Manufacturer Name')),
                item['Date'],
                item['Site'],
                str(item.get('POS no.', '101')).strip(),
                offer_str,
            )
            pair_groups.setdefault(key, {'offer': offer_str, 'parsed': parsed, 'items': []})
            pair_groups[key]['items'].append(item)

        for group in pair_groups.values():
            offer_str = group['offer']
            pair = unpack_pair_offer(group['parsed'])
            if not pair:
                pending.extend(group['items'])
                continue
            chunk_size, paid_count, discount_val, is_percentage = pair
            items = group['items']
            items.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)

            k = len(items) // chunk_size
            orphans = items[k * chunk_size:]
            for b_idx in range(k):
                bundle = items[b_idx * chunk_size: (b_idx + 1) * chunk_size]
                add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage)

            while len(orphans) >= 2:
                bundle = orphans[:2]
                orphans = orphans[2:]
                add_pair_bundle(offer_str, bundle)

            # Keep only Skip / invalid-offer singles; drop offer-mapped orphans
            for item in orphans:
                offer_for_item = resolve_offer(mapping_lookup, item.get('Manufacturer Name'), 'Skip')
                parsed_item = parse_offer_string(offer_for_item)
                if offer_for_item == 'Skip' or not parsed_item:
                    pending.append(item)

        true_unprocessed = pending

    def expand_transaction_items(trans_group):
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
        return trans_items

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

        if is_direct_offer(parsed_offer):
            process_direct_discount(group, offer_str, parsed_offer[1])
            continue

        pair = unpack_pair_offer(parsed_offer)
        if not pair:
            true_unprocessed.extend(group.to_dict('records'))
            continue
        chunk_size, paid_count, discount_val, is_percentage = pair
        company_leftovers = []

        # Phase 1: pair inside each transaction first
        for (d_val, s_val, t_val), trans_group in group.groupby(['Date', 'Site', 'Trnsctn number']):
            trans_items = expand_transaction_items(trans_group)
            trans_items.sort(key=lambda x: float(x['Gross Sales EGP']), reverse=True)
            k = len(trans_items) // chunk_size
            company_leftovers.extend(trans_items[k * chunk_size:])
            for b_idx in range(k):
                bundle = trans_items[b_idx * chunk_size: (b_idx + 1) * chunk_size]
                add_bundle(offer_str, bundle, paid_count, discount_val, is_percentage)

        # Phase 2: pair leftovers on same receipt / same POS register
        process_leftover_pool(
            company_leftovers, offer_str, chunk_size, paid_count, discount_val, is_percentage
        )

    sweep_unprocessed_receipt_pairs()

    true_unprocessed = _filter_unprocessed_export(true_unprocessed, processed_data)

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
            df_unp = pd.DataFrame(_filter_unprocessed_export(true_unprocessed, processed_data))
            if not df_unp.empty:
                final_cols_unp = [col for col in TEMPLATE_COLUMNS if col in df_unp.columns]
                df_unp[final_cols_unp].to_excel(writer, index=False, sheet_name='Unprocessed')

    buffer.seek(0)
    return buffer.getvalue()
