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
from hh_utils import DEFAULT_SETTINGS, DEFAULT_SILENCE_THRESHOLD, normalize_sku, normalize_location
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


def determine_status_vectorized(row, rules_can, rules_acc, location=None):
    """Determine inventory status with location context."""
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
        
        # Validate inputs
        from hh_utils import validate_column_mapping
        
        rules_cannabis = settings.get('cannabis_logic', DEFAULT_SETTINGS['cannabis_logic'])
        rules_accessory = settings.get('accessory_logic', DEFAULT_SETTINGS['accessory_logic'])
        col_map = settings.get('column_mapping', DEFAULT_SETTINGS['column_mapping'])
        
        required_cols = ['sku', 'qty_sold']
        
        # Check that mappings exist
        for req in required_cols:
            if req not in col_map or not col_map[req]:
                log_func(f"❌ FATAL: Column mapping missing '{req}'")
                finished_callback(False)
                return
        
        log_func("✅ Validation passed - beginning analysis...")
        
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
                df_aglc['SKU'] = df_aglc['AGLC SKU'].apply(normalize_sku)
                df_aglc = df_aglc[df_aglc['SKU'].notna()]  # Remove invalid SKUs
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
        
        # FIXED: Use consistent location normalization
        if 'Location' in df_sales.columns:
            df_sales['Loc_Key'] = df_sales['Location'].apply(normalize_location)
            
            # Log unmapped locations
            unmapped = df_sales[df_sales['Loc_Key'].str.contains('UNMAPPED', na=False)]
            if len(unmapped) > 0:
                log_func(f"⚠️  {len(unmapped)} sales records have unmapped locations:")
                log_func(f"   Examples: {unmapped['Location'].unique()[:5]}")
        else:
            log_func("⚠️  No Location column in sales data - treating as 'Other'")
            df_sales['Loc_Key'] = 'Other'
            
        pivot_sales = df_sales.pivot_table(
            index=col_map['sku'], columns='Loc_Key',
            values=[col_map['qty_sold']],
            aggfunc='sum'
        ).fillna(0)
        pivot_sales.columns = [f"{c[0]}_{c[1]}" for c in pivot_sales.columns]
        
        # === ADD: Create pivot tables for ALL sales metrics per location ===
        # Dictionary to store all pivoted metrics
        sales_pivots = {}
        
        # Metrics to capture: quantity, profit, net sales, gross sales
        metric_columns = {
            'qty_sold': col_map.get('qty_sold', 'Quantity'),
            'profit': col_map.get('profit', 'Profit'),
            'net_sales': col_map.get('net_sales', 'Net sales'),
            'gross_sales': col_map.get('gross_sales', 'Gross sales')
        }
        
        # Create pivot for each metric
        for metric_key, metric_col in metric_columns.items():
            if metric_col in df_sales.columns:
                pivot_temp = df_sales.pivot_table(
                    index=col_map['sku'],
                    columns='Loc_Key',
                    values=metric_col,
                    aggfunc='sum'
                ).fillna(0)
                
                # Store each location's data with metric in column name
                for loc in ['Hill', 'Valley', 'Jasper']:
                    if loc in pivot_temp.columns:
                        col_name = f"{metric_key}_{loc}"
                        sales_pivots[col_name] = pivot_temp[loc]
        
        # Convert to DataFrame for merging
        sales_metrics_df = pd.DataFrame(sales_pivots)
        
        # === 3. MERGE ===
        # FIXED: Use consistent SKU normalization
        master = df_inv.copy()
        if col_map['inventory_sku'] != 'SKU':
            master.rename(columns={col_map['inventory_sku']: 'SKU'}, inplace=True)
        master['SKU'] = master['SKU'].apply(normalize_sku)
        invalid_skus = master[master['SKU'].isna()]
        if len(invalid_skus) > 0:
            log_func(f"⚠️  {len(invalid_skus)} rows have invalid SKUs")
            master = master[master['SKU'].notna()]
        
        master = pd.merge(master, pivot_sales, left_on='SKU', right_index=True, how='left').fillna(0)
        
        # === Merge all sales metrics into master ===
        if not sales_metrics_df.empty:
            master = master.merge(
                sales_metrics_df,
                left_on='SKU',
                right_index=True,
                how='left'
            ).fillna(0)
            log_func(f"✓ Added {len(sales_metrics_df.columns)} sales metric columns")
        
        master = pd.merge(master, df_aglc, on='SKU', how='left')
        master['Case_Size'] = master['Case_Size'].fillna(1)
        
        # Transfers
        def prep_transfer(df, name):
            if df.empty: return pd.Series(dtype=float, name=name)
            sku_c = 'SKU' if 'SKU' in df.columns else col_map['sku']
            qty_c = 'Quantity' if 'Quantity' in df.columns else col_map['qty_sold']
            if sku_c in df.columns and qty_c in df.columns:
                df[sku_c] = df[sku_c].apply(normalize_sku)
                df = df[df[sku_c].notna()]  # Remove invalid SKUs
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
            master[f'{loc}_Status'] = temp_df.apply(lambda r: determine_status_vectorized(r, rules_cannabis, rules_accessory, location=loc), axis=1)
            master[f'{loc}_Sold'] = master[col_sold]
        
        # === 6. EXPORT - WITH EXCEL OUTLINE GROUPING (EXACTLY AS SHOWN) ===
        
        output_filename = f'Order_Rec_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
        log_func(f"Writing {output_filename}...")
        
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        workbook = writer.book
        sheet = workbook.add_worksheet("Order Builder")
        
        # === FORMAT DEFINITIONS ===
        fmt_head = workbook.add_format({
            'bold': True,
            'bg_color': '#2F5496',
            'font_color': '#FFFFFF',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        fmt_norm = workbook.add_format({'border': 1, 'align': 'left'})
        fmt_dec = workbook.add_format({'num_format': '0.00', 'border': 1, 'align': 'right'})
        fmt_currency = workbook.add_format({'num_format': '$#,##0.00', 'border': 1, 'align': 'right'})
        fmt_buy = workbook.add_format({
            'bold': True,
            'bg_color': '#FFFF00',
            'border': 1,
            'align': 'center'
        })
        
        # Product columns and all metrics per location
        product_cols = ['SKU', 'Product Name', 'Category', 'Brand', 'Case_Size']
        locations = ['Hill', 'Valley', 'Jasper']
        metrics_per_location = ['Status', 'Buy(Cs)', 'Inc', 'Sold', 'Profit', 'Net Sales', 'Gross Sales', 'Stock', 'Vel', 'WOS']
        if history_available:
            metrics_per_location.append('Trend')
        
        # === SET COLUMN WIDTHS ===
        sheet.set_column('A:A', 12)   # SKU
        sheet.set_column('B:B', 22)   # Product Name
        sheet.set_column('C:C', 12)   # Category
        sheet.set_column('D:D', 12)   # Brand
        sheet.set_column('E:E', 10)   # Case_Size
        
        # Metric columns - auto width
        for col_offset in range(len(locations) * len(metrics_per_location)):
            col_num = 5 + col_offset
            sheet.set_column(col_num, col_num, 11)
        
        # === WRITE HEADER ROW ===
        col_idx = 0
        
        # Product columns
        for h in product_cols:
            sheet.write(0, col_idx, h, fmt_head)
            col_idx += 1
        
        # Metrics headers - ONE ROW showing all metrics for all locations
        for loc in locations:
            for metric in metrics_per_location:
                sheet.write(0, col_idx, f"{loc} {metric}", fmt_head)
                col_idx += 1
        
        # === WRITE DATA ROWS ===
        for r, row in master.iterrows():
            xls_r = r + 1  # Start at row 1 (after header)
            col_idx = 0
            
            # Product columns
            for col in product_cols:
                val = row.get(col, "")
                sheet.write(xls_r, col_idx, val, fmt_norm)
                col_idx += 1
            
            # ONE ROW - all location data left to right
            for loc in locations:
                # Status
                status = row.get(f"{loc}_Status", "-")
                sheet.write(xls_r, col_idx, status, fmt_norm)
                col_idx += 1
                
                # Buy(Cs)
                buy_val = row.get(f"{loc}_SOQ", 0)
                if buy_val > 0:
                    sheet.write(xls_r, col_idx, buy_val, fmt_buy)
                else:
                    sheet.write(xls_r, col_idx, "-", fmt_norm)
                col_idx += 1
                
                # Inc
                inc_str = row.get(f"{loc}_Inc_Str", "-")
                sheet.write(xls_r, col_idx, inc_str, fmt_norm)
                col_idx += 1
                
                # Sold
                sold = row.get(f"qty_sold_{loc}", row.get(f"{loc}_Sold", 0))
                sheet.write(xls_r, col_idx, sold, fmt_dec)
                col_idx += 1
                
                # Profit
                profit = row.get(f"profit_{loc}", 0)
                sheet.write(xls_r, col_idx, profit, fmt_currency)
                col_idx += 1
                
                # Net Sales
                net_sales = row.get(f"net_sales_{loc}", 0)
                sheet.write(xls_r, col_idx, net_sales, fmt_currency)
                col_idx += 1
                
                # Gross Sales
                gross_sales = row.get(f"gross_sales_{loc}", 0)
                sheet.write(xls_r, col_idx, gross_sales, fmt_currency)
                col_idx += 1
                
                # Stock
                stock = row.get(f"{loc}_Stock", 0)
                sheet.write(xls_r, col_idx, stock, fmt_dec)
                col_idx += 1
                
                # Velocity
                vel = row.get(f"{loc}_Vel", 0)
                sheet.write(xls_r, col_idx, vel, fmt_dec)
                col_idx += 1
                
                # WOS
                wos = row.get(f"{loc}_WOS", 0)
                if wos >= DEFAULT_SILENCE_THRESHOLD:
                    sheet.write(xls_r, col_idx, "∞", fmt_norm)
                else:
                    sheet.write(xls_r, col_idx, wos, fmt_dec)
                col_idx += 1
                
                # Trend (if available)
                if history_available:
                    trend = row.get(f"{loc}_Trend", "-")
                    sheet.write(xls_r, col_idx, trend, fmt_norm)
                    col_idx += 1
        
        # === SET UP OUTLINE GROUPING ===
        # This creates the +/- collapse buttons on the left margin
        
        product_col_count = len(product_cols)
        metrics_count = len(metrics_per_location)
        
        # For each location, group its columns together
        # Set outline levels: Status is level 1 (always visible), other metrics are level 2 (collapsible)
        for loc_num, loc in enumerate(locations):
            group_start = product_col_count + (loc_num * metrics_count)
            
            # Status column (always visible when collapsed) - level 1
            sheet.set_column(group_start, group_start, None, None, {'level': 1, 'collapsed': False})
            
            # All other metrics for this location - level 2 (can be collapsed)
            for metric_idx in range(1, metrics_count):
                col_num = group_start + metric_idx
                sheet.set_column(col_num, col_num, None, None, {'level': 2, 'collapsed': False})
        
        # Set outline display - show outline controls, default to level 1 (shows Status only when collapsed)
        sheet.outline_settings(True, False, False, True)
        
        # Freeze panes (keep product columns visible)
        sheet.freeze_panes(1, len(product_cols))
        
        writer.close()
        
        try:
            os.startfile(output_filename)
        except:
            subprocess.call(['open', output_filename])
        
        log_func(f"✅ Report saved: {output_filename}")
        log_func(f"   Tip: Use +/- buttons on LEFT to collapse/expand each location group!")
        finished_callback(True)
        
    except Exception as e:
        log_func(f"❌ FATAL: {e}")
        traceback.print_exc()
        finished_callback(False)