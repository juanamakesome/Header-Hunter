"""
Header Hunter v5.0 - Business Logic Module
Core data processing and inventory status determination
"""
import pandas as pd
import numpy as np
import xlsxwriter
import openpyxl
import os
import subprocess
import re
import traceback
from datetime import datetime
from hh_utils import DEFAULT_SETTINGS, DEFAULT_SILENCE_THRESHOLD


def clean_currency(val):
    """
    Convert currency values to float, handling various formats.
    Supports parenthetical negatives (1,234.56) and standard formats.
    
    Args:
        val: Value to clean (any type)
        
    Returns:
        float: Cleaned numeric value, 0.0 if conversion fails
    """
    if pd.isna(val):
        return 0.0
    
    val_str = str(val)
    
    # Handle parenthetical negatives: (1234.56) -> -1234.56
    if val_str.startswith('(') and val_str.endswith(')'):
        val_str = '-' + val_str[1:-1]
    
    # Remove all non-numeric characters except decimal and negative sign
    clean = re.sub(r'[^\d.-]', '', val_str)
    
    try:
        return float(clean)
    except ValueError:
        return 0.0


def determine_status_vectorized(row, rules_can, rules_acc):
    """
    Determine inventory status for a single SKU based on business logic rules.
    
    Logic hierarchy:
    1. Zero velocity items (New or Cold)
    2. High velocity items (Hot or Hot-with-alert)
    3. Medium velocity items (Good or Good-with-alert)
    4. Low velocity items (Dead or Filler)
    
    Args:
        row (pd.Series): Row with keys: Vel, WOS, Stock, Incoming_Num, Is_Accessory
        rules_can (dict): Cannabis product logic rules
        rules_acc (dict): Accessory product logic rules
        
    Returns:
        str: Status indicator with emoji (e.g., "🔥 Hot", "🚨 Reorder")
    """
    # Select appropriate rules based on product type
    rules = rules_acc if row['Is_Accessory'] else rules_can
    
    velocity = row['Vel']
    wos = row['WOS']
    effective_oh = max(0, row['Stock'])
    incoming = row['Incoming_Num']
    total_avail = effective_oh + incoming
    
    # Calculate effective WOS including incoming stock
    effective_wos = DEFAULT_SILENCE_THRESHOLD  # Default when no velocity
    if velocity > 0:
        effective_wos = total_avail / velocity
    
    # === LOGIC TREE ===
    
    # 1. ZERO VELOCITY (New or Cold)
    if velocity == 0 and (incoming > 0 or effective_oh > 0):
        if incoming > 0:
            return "✨ New"    # Incoming stock, no demand yet
        return "❄️ Cold"       # Stocked but no sales
    
    # 2. HIGH VELOCITY (Hot)
    if velocity >= rules['hot_velocity']:
        if wos < rules['reorder_point']:
            if effective_wos >= rules['reorder_point']:
                return "🚚 Landing"   # Incoming will cover
            return "🚨 Reorder"       # Critical: need to reorder
        return "🔥 Hot"               # Strong sales, adequate stock
    
    # 3. MEDIUM VELOCITY (Good)
    # Threshold is 25% of hot velocity
    good_vel_threshold = rules['hot_velocity'] * 0.25
    if velocity >= good_vel_threshold:
        if wos < rules['reorder_point']:
            if effective_wos >= rules['reorder_point']:
                return "🚚 Landing"
            return "🚨 Reorder"
        return "✅ Good"              # Steady sales, adequate stock
    
    # 4. LOW VELOCITY (Dead)
    if wos > rules['dead_wos'] and effective_oh > rules['dead_on_hand']:
        return "💀 Dead"              # High stock, minimal sales
    
    # Default: Minimal stock, minimal velocity
    return "➖"


def run_logic_pandas(file_paths, settings, report_days, log_func, finished_callback):
    """
    Main analysis engine. Orchestrates data loading, processing, and report generation.
    
    Workflow:
    1. Load data from CSV/Excel files
    2. Clean and normalize column names
    3. Merge datasets and calculate metrics
    4. Determine status for each SKU per location
    5. Generate Excel workbook with recommendations
    
    Args:
        file_paths (dict): File paths keyed by: inventory, sales, po, aglc, hill, valley, jasper
        settings (dict): Configuration including logic rules and column mappings
        report_days (str): Number of days of sales data to analyze
        log_func (callable): Callback for logging messages
        finished_callback (callable): Callback when complete (bool success parameter)
    """
    try:
        log_func("--- Starting Analysis (Modular v5.0) ---")
        
        rules_cannabis = settings.get('cannabis_logic', DEFAULT_SETTINGS['cannabis_logic'])
        rules_accessory = settings.get('accessory_logic', DEFAULT_SETTINGS['accessory_logic'])
        col_map = settings.get('column_mapping', DEFAULT_SETTINGS['column_mapping'])
        
        # Parse report days with validation
        try:
            days = float(report_days)
            if days <= 0:
                days = 30
        except (ValueError, TypeError):
            days = 30
        
        weeks_factor = days / 7.0
        log_func("Loading Data Files...")
        
        # === 1. READ DATA ===
        try:
            df_inv = pd.read_csv(file_paths['inventory'])
            df_sales = pd.read_csv(file_paths['sales'])
            
            def load_opt(path, cols):
                """Load optional file or return empty DataFrame."""
                return pd.read_csv(path) if path else pd.DataFrame(columns=cols)
            
            df_po = load_opt(file_paths['po'], [col_map['sku'], 'Quantity ordered'])
            df_hill = load_opt(file_paths['hill'], [col_map['sku'], 'Quantity'])
            df_valley = load_opt(file_paths['valley'], [col_map['sku'], 'Quantity'])
            df_jasper = load_opt(file_paths['jasper'], [col_map['sku'], 'Quantity'])
            
            # AGLC manual form (special parsing)
            if file_paths['aglc']:
                log_func(" -> Reading AGLC Manual Order Form...")
                df_aglc = pd.read_excel(file_paths['aglc'], header=10, engine='openpyxl')
                df_aglc['SKU'] = df_aglc['AGLC SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_aglc['Case_Size'] = pd.to_numeric(df_aglc['EachesPerCase'], errors='coerce').fillna(1)
                df_aglc['Case_Cost'] = pd.to_numeric(df_aglc['Sell Price Per Case'], errors='coerce').fillna(0)
                df_aglc = df_aglc[['SKU', 'Case_Size', 'Case_Cost']].drop_duplicates(subset=['SKU'])
            else:
                df_aglc = pd.DataFrame(columns=['SKU', 'Case_Size', 'Case_Cost'])
                
        except Exception as e:
            log_func(f"❌ Error loading files: {e}")
            finished_callback(False)
            return
        
        log_func("Processing Sales Data...")
        
        # === 2. CLEAN & PIVOT SALES ===
        cols_to_clean = [col_map['qty_sold'], col_map['profit'], col_map['net_sales'], col_map['gross_sales']]
        for col in cols_to_clean:
            if col in df_sales.columns:
                df_sales[col] = df_sales[col].apply(clean_currency)
        
        def normalize_loc(loc):
            """Normalize location names to standard format."""
            s_loc = str(loc)
            if 'Hill' in s_loc:
                return 'Hill'
            if 'Valley' in s_loc:
                return 'Valley'
            if 'Jasper' in s_loc:
                return 'Jasper'
            return 'Other'
        
        # Create location key column
        if 'Location' in df_sales.columns:
            df_sales['Loc_Key'] = df_sales['Location'].apply(normalize_loc)
        else:
            df_sales['Loc_Key'] = 'Other'
        
        # Pivot sales data by location
        pivot_sales = df_sales.pivot_table(
            index=col_map['sku'],
            columns='Loc_Key',
            values=[col_map['qty_sold'], col_map['gross_sales'], col_map['net_sales'], col_map['profit']],
            aggfunc='sum'
        ).fillna(0)
        
        pivot_sales.columns = [f"{c[0]}_{c[1]}" for c in pivot_sales.columns]
        
        log_func("Merging Data...")
        
        # === 3. MERGE DATASETS ===
        master = df_inv.copy()
        
        # Normalize SKU column name
        if col_map['inventory_sku'] != 'SKU':
            master.rename(columns={col_map['inventory_sku']: 'SKU'}, inplace=True)
        
        # Clean SKU format (remove .0 decimals)
        master['SKU'] = master['SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
        
        # Merge with sales pivot, AGLC data
        master = pd.merge(master, pivot_sales, left_on='SKU', right_index=True, how='left').fillna(0)
        master = pd.merge(master, df_aglc, on='SKU', how='left')
        master['Case_Size'] = master['Case_Size'].fillna(1)
        master['Case_Cost'] = master['Case_Cost'].fillna(0)
        
        # === 4. PROCESS TRANSFERS ===
        def prep_transfer(df, name):
            """Prepare transfer/PO data: sum by SKU."""
            if df.empty:
                return pd.Series(dtype=float, name=name)
            
            sku_c = col_map['sku'] if col_map['sku'] in df.columns else 'SKU'
            qty_c = col_map['qty_sold'] if col_map['qty_sold'] in df.columns else 'Quantity'
            
            if sku_c in df.columns and qty_c in df.columns:
                df[sku_c] = df[sku_c].astype(str).str.replace(r'\.0$', '', regex=True)
                df[qty_c] = df[qty_c].apply(clean_currency)
                return df.groupby(sku_c)[qty_c].sum().rename(name)
            
            return pd.Series(dtype=float, name=name)
        
        t_hill = prep_transfer(df_hill, 'Trans_Hill')
        t_valley = prep_transfer(df_valley, 'Trans_Valley')
        t_jasper = prep_transfer(df_jasper, 'Trans_Jasper')
        
        if not df_po.empty:
            sku_c = col_map['sku'] if col_map['sku'] in df_po.columns else 'SKU'
            df_po[sku_c] = df_po[sku_c].astype(str).str.replace(r'\.0$', '', regex=True)
            df_po['Quantity ordered'] = df_po['Quantity ordered'].apply(clean_currency)
            po_series = df_po.groupby(sku_c)['Quantity ordered'].sum().rename('PO_Qty')
        else:
            po_series = pd.Series(dtype=float, name='PO_Qty')
        
        # Join transfer data
        master = master.set_index('SKU')
        master = master.join([t_hill, t_valley, t_jasper, po_series], how='left').fillna(0)
        master = master.reset_index()
        
        log_func("Running Algorithms (Vectorized)...")
        
        # === 5. CALCULATE METRICS PER LOCATION ===
        master['Is_Accessory'] = ~master['SKU'].astype(str).str.upper().str.startswith("CNB-")
        master['Target_WOS'] = np.where(
            master['Is_Accessory'],
            rules_accessory['target_wos'],
            rules_cannabis['target_wos']
        )
        
        # Calculate net PO for Jasper (after transfers to other locations)
        master['PO_Net_Jasper'] = (master['PO_Qty'] - master['Trans_Hill'] - master['Trans_Valley']).clip(lower=0)
        
        locations = ['Hill', 'Valley', 'Jasper']
        q_key = col_map['qty_sold']
        rev_key = col_map['net_sales']
        gross_key = col_map['gross_sales']
        prof_key = col_map['profit']
        
        # Process each location
        for loc in locations:
            col_sold = f'{q_key}_{loc}'
            col_rev = f'{rev_key}_{loc}'
            col_gross = f'{gross_key}_{loc}'
            col_prof = f'{prof_key}_{loc}'
            
            # Ensure all required columns exist
            for c in [col_sold, col_rev, col_gross, col_prof]:
                if c not in master.columns:
                    master[c] = 0.0
            
            # Find and sum stock columns for this location
            stock_cols = [
                c for c in master.columns
                if loc in c and
                ('Sales' in c or 'Storage' in c or 'Inventory' in c) and
                c not in [col_sold, col_rev, col_gross, col_prof]
            ]
            
            for sc in stock_cols:
                if master[sc].dtype == object:
                    master[sc] = master[sc].apply(clean_currency)
            
            master[f'{loc}_Stock'] = master[stock_cols].sum(axis=1)
            
            # Set incoming quantities based on transfer source
            if loc == 'Hill':
                master[f'{loc}_Inc_Num'] = master['Trans_Hill']
                master[f'{loc}_Inc_Str'] = np.where(
                    master['Trans_Hill'] > 0,
                    master['Trans_Hill'].astype(int).astype(str) + " 🚚",
                    "-"
                )
            elif loc == 'Valley':
                master[f'{loc}_Inc_Num'] = master['Trans_Valley']
                master[f'{loc}_Inc_Str'] = np.where(
                    master['Trans_Valley'] > 0,
                    master['Trans_Valley'].astype(int).astype(str) + " 🚚",
                    "-"
                )
            elif loc == 'Jasper':
                master[f'{loc}_Inc_Num'] = master['PO_Net_Jasper']
                master[f'{loc}_Inc_Str'] = np.where(
                    master['PO_Net_Jasper'] > 0,
                    master['PO_Net_Jasper'].astype(int).astype(str) + " 📦",
                    "-"
                )
            
            # Calculate velocity (units per week)
            master[f'{loc}_Vel'] = master[col_sold] / weeks_factor
            
            # Calculate weeks on stock
            master[f'{loc}_WOS'] = np.where(
                master[f'{loc}_Vel'] > 0,
                master[f'{loc}_Stock'] / master[f'{loc}_Vel'],
                np.where(master[f'{loc}_Stock'] > 0, DEFAULT_SILENCE_THRESHOLD, 0.0)
            )
            
            # Calculate margin percentage
            master[f'{loc}_Mrg'] = np.where(
                master[col_rev] > 0,
                master[col_prof] / master[col_rev],
                0.0
            )
            
            # Calculate order quantity
            target_stock = master[f'{loc}_Vel'] * master['Target_WOS']
            net_need = target_stock - (master[f'{loc}_Stock'] + master[f'{loc}_Inc_Num'])
            safe_case_size = np.maximum(master['Case_Size'], 1)
            master[f'{loc}_SOQ'] = np.ceil(np.maximum(net_need, 0) / safe_case_size)
            
            # Determine status
            temp_df = pd.DataFrame({
                'Vel': master[f'{loc}_Vel'],
                'WOS': master[f'{loc}_WOS'],
                'Stock': master[f'{loc}_Stock'],
                'Incoming_Num': master[f'{loc}_Inc_Num'],
                'Is_Accessory': master['Is_Accessory']
            })
            
            master[f'{loc}_Status'] = temp_df.apply(
                lambda r: determine_status_vectorized(r, rules_cannabis, rules_accessory),
                axis=1
            )
            
            # Copy financial data
            master[f'{loc}_Sold'] = master[col_sold]
            master[f'{loc}_Gross'] = master[col_gross]
            master[f'{loc}_Profit'] = master[col_prof]
        
        # === 6. WRITE EXCEL REPORT ===
        output_filename = f'Order_Rec_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
        log_func(f"Writing Excel: {output_filename}...")
        
        cols_static = ['SKU', 'Product Name', 'Category', 'Brand', 'Case_Size', 'Case_Cost']
        
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        workbook = writer.book
        
        # Define formats
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        fmt_text = workbook.add_format({'border': 1})
        fmt_num = workbook.add_format({'border': 1})
        fmt_dec = workbook.add_format({'num_format': '0.00', 'border': 1})
        fmt_pct = workbook.add_format({'num_format': '0.0%', 'border': 1})
        fmt_curr = workbook.add_format({'num_format': '$#,##0.00', 'border': 1})
        fmt_inc = workbook.add_format({'border': 1, 'align': 'center'})
        fmt_soq = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        
        # Location-specific header formats
        loc_configs = [
            ('Hill', workbook.add_format({'bold': True, 'bg_color': '#B4C6E7', 'border': 1})),
            ('Valley', workbook.add_format({'bold': True, 'bg_color': '#F8CBAD', 'border': 1})),
            ('Jasper', workbook.add_format({'bold': True, 'bg_color': '#C6E0B4', 'border': 1}))
        ]
        
        worksheet = workbook.add_worksheet("Order Builder")
        writer.sheets['Order Builder'] = worksheet
        
        # Write header row
        current_col = 0
        for h in ["Key"] + cols_static:
            worksheet.write(0, current_col, h, fmt_header)
            current_col += 1
        
        # Metrics to display per location
        metrics = [
            ('Status', 'Status', fmt_text),
            ('Buy(Cs)', 'SOQ', fmt_soq),
            ('Incoming', 'Inc_Str', fmt_inc),
            ('Sold', 'Sold', fmt_num),
            ('Gross', 'Gross', fmt_curr),
            ('Profit', 'Profit', fmt_curr),
            ('Mrg%', 'Mrg', fmt_pct),
            ('Stock', 'Stock', fmt_num),
            ('Vel', 'Vel', fmt_dec),
            ('WOS', 'WOS', fmt_dec)
        ]
        
        for loc_name, fmt in loc_configs:
            for title, _, _ in metrics:
                worksheet.write(0, current_col, f"{loc_name} {title}", fmt)
                current_col += 1
        
        # Write data rows
        for r_idx, row in master.iterrows():
            c_idx = 0
            xls_row = r_idx + 1
            
            # Index column
            worksheet.write(xls_row, c_idx, "", fmt_text)
            c_idx += 1
            
            # Static columns
            for col in cols_static:
                val = row.get(col, "")
                if pd.isna(val):
                    val = ""
                
                if col == 'Case_Cost':
                    worksheet.write(xls_row, c_idx, val, fmt_curr)
                else:
                    worksheet.write(xls_row, c_idx, val, fmt_text)
                c_idx += 1
            
            # Location metrics
            for loc_name, _ in loc_configs:
                for m_idx, (_, key_suffix, cell_fmt) in enumerate(metrics):
                    val = row.get(f"{loc_name}_{key_suffix}", 0)
                    
                    if key_suffix == 'SOQ' and (pd.isna(val) or val == 0):
                        worksheet.write(xls_row, c_idx, "-", cell_fmt)
                    elif pd.isna(val):
                        worksheet.write(xls_row, c_idx, 0, cell_fmt)
                    else:
                        worksheet.write(xls_row, c_idx, val, cell_fmt)
                    
                    if r_idx == 0:
                        if m_idx >= 2:
                            worksheet.set_column(c_idx, c_idx, 10, None, {'level': 1, 'hidden': False})
                        else:
                            worksheet.set_column(c_idx, c_idx, 12, None, {'level': 0})
                    
                    c_idx += 1
        
        # Freeze header and key columns
        worksheet.freeze_panes(1, 7)
        
        writer.close()
        
        log_func("✅ Success! Report Generated.")
        
        # === 7. OPEN REPORT ===
        try:
            # Windows
            os.startfile(output_filename)
        except AttributeError:
            # macOS/Linux
            subprocess.call(['open', output_filename])
        
        finished_callback(True)
        
    except Exception as e:
        log_func(f"❌ FATAL ERROR: {e}")
        traceback.print_exc()
        finished_callback(False)