"""
Header Hunter - Business Logic Module
Fixed: Separate Accessory Buying Logic and Sales Data Integration
"""
import pandas as pd
import numpy as np
import xlsxwriter
import os
import subprocess
import re
import traceback
from datetime import datetime
from hh_utils import DEFAULT_SETTINGS, DEFAULT_SILENCE_THRESHOLD
# UPDATED IMPORT:
from hh_history import load_history_db, compute_rolling_velocities


def clean_currency(val):
    if pd.isna(val): 
        return 0.0
    val_str = str(val)
    if val_str.startswith('(') and val_str.endswith(')'): 
        val_str = '-' + val_str[1:-1]
    clean = re.sub(r'[^\d.-]', '', val_str)
    try: 
        return float(clean)
    except: 
        return 0.0


def determine_status_vectorized(row, rules_can, rules_acc):
    """Determine inventory status."""
    is_acc = row['Is_Accessory']
    rules = rules_acc if is_acc else rules_can
    
    velocity = row['Vel']
    wos = row['WOS']
    effective_oh = max(0, row['Stock'])
    incoming = row['Incoming_Num']
    
    # 1. Accessories Logic (Simplified)
    if is_acc:
        if velocity == 0 and effective_oh == 0: 
            return "➖"
        if velocity == 0 and effective_oh > 0: 
            return "❄️ Cold"
        if velocity > rules['hot_velocity']: 
            return "🔥 Hot"
        if wos < rules['reorder_point']: 
            return "⚠️ Low" 
        return "✅ Good"

    # 2. Cannabis Logic
    if velocity == 0:
        if incoming > 0: return "✨ New"
        if effective_oh > 0: return "❄️ Cold"
        return "➖"
    
    if velocity >= rules['hot_velocity']:
        if wos < rules['reorder_point']:
            if (effective_oh + incoming) / velocity >= rules['reorder_point']:
                return "🚚 Landing"
            return "🚨 Reorder"
        return "🔥 Hot"
    
    good_threshold = rules['hot_velocity'] * 0.25
    if velocity >= good_threshold:
        if wos < rules['reorder_point']:
            return "🚨 Reorder"
        return "✅ Good"
    
    if wos > rules['dead_wos'] and effective_oh > rules['dead_on_hand']:
        return "💀 Dead"
        
    return "➖"


def run_logic_pandas(file_paths, settings, report_days, log_func, finished_callback):
    """Main analysis engine."""
    try:
        log_func("--- Starting Analysis (Fixed v5.2) ---")
        
        rules_cannabis = settings.get('cannabis_logic', DEFAULT_SETTINGS['cannabis_logic'])
        rules_accessory = settings.get('accessory_logic', DEFAULT_SETTINGS['accessory_logic'])
        col_map = settings.get('column_mapping', DEFAULT_SETTINGS['column_mapping'])
        
        try: days = float(report_days)
        except: days = 30
        if days <= 0: days = 30
        weeks_factor = days / 7.0
        
        # === 1. READ DATA ===
        log_func("Loading Data Files...")
        df_inv = pd.read_csv(file_paths['inventory'])
        df_sales = pd.read_csv(file_paths['sales'])
        
        def load_opt(path): 
            return pd.read_csv(path) if path else pd.DataFrame()
        df_po = load_opt(file_paths['po'])
        df_hill = load_opt(file_paths['hill'])
        df_valley = load_opt(file_paths['valley'])
        df_jasper = load_opt(file_paths['jasper'])
        
        # Load AGLC
        if file_paths['aglc']:
            try:
                df_aglc = pd.read_excel(file_paths['aglc'], header=10, engine='openpyxl')
                df_aglc['SKU'] = df_aglc['AGLC SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
                df_aglc['Case_Size'] = pd.to_numeric(df_aglc['EachesPerCase'], errors='coerce').fillna(1)
                df_aglc = df_aglc[['SKU', 'Case_Size']].drop_duplicates(subset=['SKU'])
            except:
                df_aglc = pd.DataFrame(columns=['SKU', 'Case_Size'])
        else:
            df_aglc = pd.DataFrame(columns=['SKU', 'Case_Size'])
        
        # === 2. CLEAN SALES ===
        cols_to_clean = [col_map['qty_sold'], col_map['profit'], col_map['net_sales'], col_map['gross_sales']]
        for col in cols_to_clean:
            if col in df_sales.columns:
                df_sales[col] = df_sales[col].apply(clean_currency)
        
        if 'Location' in df_sales.columns:
            def normalize_loc(loc):
                s_loc = str(loc)
                if 'Hill' in s_loc: return 'Hill'
                if 'Valley' in s_loc: return 'Valley'
                if 'Jasper' in s_loc: return 'Jasper'
                return 'Other'
            df_sales['Loc_Key'] = df_sales['Location'].apply(normalize_loc)
        else:
            df_sales['Loc_Key'] = 'Other'
            
        pivot_sales = df_sales.pivot_table(
            index=col_map['sku'], columns='Loc_Key',
            values=[col_map['qty_sold']],
            aggfunc='sum'
        ).fillna(0)
        pivot_sales.columns = [f"{c[0]}_{c[1]}" for c in pivot_sales.columns]
        
        # === 3. MERGE ===
        master = df_inv.copy()
        if col_map['inventory_sku'] != 'SKU':
            master.rename(columns={col_map['inventory_sku']: 'SKU'}, inplace=True)
        master['SKU'] = master['SKU'].astype(str).str.replace(r'\.0$', '', regex=True).str.upper()
        
        master = pd.merge(master, pivot_sales, left_on='SKU', right_index=True, how='left').fillna(0)
        master = pd.merge(master, df_aglc, on='SKU', how='left')
        master['Case_Size'] = master['Case_Size'].fillna(1)
        
        # Transfers
        def prep_transfer(df, name):
            if df.empty: return pd.Series(dtype=float, name=name)
            sku_c = 'SKU' if 'SKU' in df.columns else col_map['sku']
            qty_c = 'Quantity' if 'Quantity' in df.columns else col_map['qty_sold']
            if sku_c in df.columns and qty_c in df.columns:
                df[sku_c] = df[sku_c].astype(str).str.replace(r'\.0$', '', regex=True).str.upper()
                df[qty_c] = df[qty_c].apply(clean_currency)
                return df.groupby(sku_c)[qty_c].sum().rename(name)
            return pd.Series(dtype=float, name=name)

        master = master.set_index('SKU')
        master = master.join([
            prep_transfer(df_hill, 'Trans_Hill'),
            prep_transfer(df_valley, 'Trans_Valley'),
            prep_transfer(df_jasper, 'Trans_Jasper'),
            prep_transfer(df_po, 'PO_Qty')
        ], how='left').fillna(0)
        master = master.reset_index()
        
        # === 4. HISTORY & VELOCITY (UPDATED) ===
        history_folder = settings.get('history_folder', None)
        history_available = False
        df_history = None
        
        if history_folder:
            # We now load the Master Parquet directly
            df_history = load_history_db(history_folder, log_func)
            if df_history is not None and not df_history.empty: 
                history_available = True
        
        # === 5. CALCULATE METRICS ===
        master['Is_Accessory'] = ~master['SKU'].str.startswith("CNB-")
        locations = ['Hill', 'Valley', 'Jasper']
        
        for loc in locations:
            col_sold = f'{col_map["qty_sold"]}_{loc}'
            if col_sold not in master.columns: master[col_sold] = 0.0
            
            # Stock
            stock_cols = [c for c in master.columns if loc in c and ('Sales' in c or 'Storage' in c or 'Inventory' in c) and c not in [col_sold]]
            for sc in stock_cols:
                if master[sc].dtype == object: master[sc] = master[sc].apply(clean_currency)
            master[f'{loc}_Stock'] = master[stock_cols].sum(axis=1)
            
            # Velocity: Default to current file
            master[f'{loc}_Vel'] = master[col_sold] / weeks_factor
            
            # Velocity: Override with History if available
            if history_available:
                vel_4w = []
                trends = []
                for sku in master['SKU']:
                    m = compute_rolling_velocities(df_history, sku, location=loc)
                    vel_4w.append(m['vel_4w'])
                    trends.append(m['trend'])
                
                master[f'{loc}_Vel_4w'] = vel_4w
                master[f'{loc}_Trend'] = trends
                master[f'{loc}_Vel'] = np.where(master[f'{loc}_Vel_4w'] > 0, master[f'{loc}_Vel_4w'], master[f'{loc}_Vel'])
            
            # WOS
            master[f'{loc}_WOS'] = np.where(
                master[f'{loc}_Vel'] > 0, 
                master[f'{loc}_Stock'] / master[f'{loc}_Vel'], 
                DEFAULT_SILENCE_THRESHOLD
            )
            
            # Incoming
            if loc == 'Jasper':
                net_po = (master['PO_Qty'] - master['Trans_Hill'] - master['Trans_Valley']).clip(lower=0)
                master[f'{loc}_Inc_Num'] = net_po
                master[f'{loc}_Inc_Str'] = np.where(net_po > 0, net_po.astype(int).astype(str) + " 📦", "-")
            else:
                t_col = f'Trans_{loc}'
                master[f'{loc}_Inc_Num'] = master[t_col]
                master[f'{loc}_Inc_Str'] = np.where(master[t_col] > 0, master[t_col].astype(int).astype(str) + " 🚚", "-")

            # Buy Logic
            target_wos = np.where(master['Is_Accessory'], rules_accessory['target_wos'], rules_cannabis['target_wos'])
            target_stock = master[f'{loc}_Vel'] * target_wos
            net_need = target_stock - (master[f'{loc}_Stock'] + master[f'{loc}_Inc_Num'])
            
            soq_raw = np.where(master['Is_Accessory'], 0, np.ceil(np.maximum(net_need, 0) / np.maximum(master['Case_Size'], 1)))
            master[f'{loc}_SOQ'] = soq_raw
            
            # Status
            temp_df = pd.DataFrame({
                'Vel': master[f'{loc}_Vel'],
                'WOS': master[f'{loc}_WOS'],
                'Stock': master[f'{loc}_Stock'],
                'Incoming_Num': master[f'{loc}_Inc_Num'],
                'Is_Accessory': master['Is_Accessory']
            })
            master[f'{loc}_Status'] = temp_df.apply(lambda r: determine_status_vectorized(r, rules_cannabis, rules_accessory), axis=1)
            master[f'{loc}_Sold'] = master[col_sold]
        
        # === 6. EXPORT ===
        output_filename = f'Order_Rec_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
        log_func(f"Writing {output_filename}...")
        
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        workbook = writer.book
        sheet = workbook.add_worksheet("Order Builder")
        
        fmt_head = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        fmt_norm = workbook.add_format({'border': 1})
        fmt_dec = workbook.add_format({'num_format': '0.00', 'border': 1})
        fmt_buy = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        
        cols = ['SKU', 'Product Name', 'Category', 'Brand', 'Case_Size']
        locs = [('Hill', '#B4C6E7'), ('Valley', '#F8CBAD'), ('Jasper', '#C6E0B4')]
        
        c = 0
        for h in ["Key"] + cols:
            sheet.write(0, c, h, fmt_head); c += 1
            
        for lname, color in locs:
            fmt_loc = workbook.add_format({'bold': True, 'bg_color': color, 'border': 1})
            metrics_list = ['Status', 'Buy(Cs)', 'Inc', 'Sold', 'Stock', 'Vel', 'WOS']
            if history_available: metrics_list.append('Trend')
            for m in metrics_list:
                sheet.write(0, c, f"{lname} {m}", fmt_loc); c += 1
                
        for r, row in master.iterrows():
            xls_r = r + 1; c = 0
            sheet.write(xls_r, c, "", fmt_norm); c += 1 # Key
            for col in cols:
                sheet.write(xls_r, c, row.get(col, ""), fmt_norm); c += 1
                
            for lname, _ in locs:
                sheet.write(xls_r, c, row.get(f"{lname}_Status", "-"), fmt_norm); c+=1
                buy_val = row.get(f"{lname}_SOQ", 0)
                if buy_val > 0: sheet.write(xls_r, c, buy_val, fmt_buy)
                else: sheet.write(xls_r, c, "-", fmt_norm)
                c+=1
                sheet.write(xls_r, c, row.get(f"{lname}_Inc_Str", "-"), fmt_norm); c+=1
                sheet.write(xls_r, c, row.get(f"{lname}_Sold", 0), fmt_norm); c+=1
                sheet.write(xls_r, c, row.get(f"{lname}_Stock", 0), fmt_norm); c+=1
                sheet.write(xls_r, c, row.get(f"{lname}_Vel", 0), fmt_dec); c+=1
                wos_val = row.get(f"{lname}_WOS", 0)
                if wos_val >= DEFAULT_SILENCE_THRESHOLD: sheet.write(xls_r, c, "∞", fmt_norm)
                else: sheet.write(xls_r, c, wos_val, fmt_dec)
                c+=1
                if history_available:
                    sheet.write(xls_r, c, row.get(f"{lname}_Trend", "-"), fmt_norm); c+=1

        sheet.freeze_panes(1, 6)
        writer.close()
        try: os.startfile(output_filename)
        except: subprocess.call(['open', output_filename])
        
        finished_callback(True)
        
    except Exception as e:
        log_func(f"❌ FATAL: {e}")
        traceback.print_exc()
        finished_callback(False)