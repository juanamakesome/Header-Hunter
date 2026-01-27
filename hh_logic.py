"""
Header Hunter v8.0 - Business Logic Module
Core data processing and inventory status determination
Built from v6.0 with full business rules integration
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
from excel_writer import write_excel_report
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




def normalize_transfer_loc(loc_str):
    """Normalize location names from transfer files to standard format."""
    if pd.isna(loc_str):
        return None
    loc_str = str(loc_str).upper()
    if 'HILL' in loc_str:
        return 'Hill'
    if 'VALLEY' in loc_str:
        return 'Valley'
    if 'JASPER' in loc_str:
        return 'Jasper'
    return None


def _process_transfer_data(df_transfers, log_func):
    """Process transfer files with Source/Dest columns."""
    if df_transfers.empty:
        return {
            'Trans_Hill': pd.Series(dtype=float, name='Trans_Hill'),
            'Trans_Valley': pd.Series(dtype=float, name='Trans_Valley'),
            'Trans_Jasper': pd.Series(dtype=float, name='Trans_Jasper'),
            'Trans_To_Hill': pd.Series(dtype=float, name='Trans_To_Hill'),
            'Trans_To_Valley': pd.Series(dtype=float, name='Trans_To_Valley'),
            'Trans_To_Jasper': pd.Series(dtype=float, name='Trans_To_Jasper')
        }
    
    # Find columns
    sku_col = None
    qty_col = None
    source_col = None
    dest_col = None
    
    for col in df_transfers.columns:
        col_upper = str(col).upper()
        if not sku_col and ('SKU' in col_upper):
            sku_col = col
        if not qty_col and ('QUANTITY' in col_upper or 'QTY' in col_upper):
            qty_col = col
        if not source_col and ('SOURCE' in col_upper):
            source_col = col
        if not dest_col and ('DEST' in col_upper or 'DESTINATION' in col_upper):
            dest_col = col
    
    if not sku_col:
        sku_col = df_transfers.columns[0]
    if not qty_col:
        qty_col = df_transfers.columns[1] if len(df_transfers.columns) > 1 else df_transfers.columns[0]
    
    # Clean SKU and Quantity
    df_transfers[sku_col] = df_transfers[sku_col].astype(str).str.replace(r'\.0$', '', regex=True)
    df_transfers[qty_col] = df_transfers[qty_col].apply(clean_currency)
    
    # Normalize Source and Dest locations
    if source_col:
        df_transfers['_Source_Norm'] = df_transfers[source_col].apply(normalize_transfer_loc)
    else:
        df_transfers['_Source_Norm'] = None
        log_func("  ⚠️ No Source column found in transfer file, transfers FROM locations will be 0")
    
    if dest_col:
        df_transfers['_Dest_Norm'] = df_transfers[dest_col].apply(normalize_transfer_loc)
    else:
        df_transfers['_Dest_Norm'] = None
        log_func("  ⚠️ No Dest column found in transfer file, transfers TO locations will be 0")
    
    # Calculate transfers FROM each location (outgoing)
    trans_from_hill = pd.Series(dtype=float, name='Trans_Hill')
    trans_from_valley = pd.Series(dtype=float, name='Trans_Valley')
    trans_from_jasper = pd.Series(dtype=float, name='Trans_Jasper')
    
    # Calculate transfers TO each location (incoming)
    trans_to_hill = pd.Series(dtype=float, name='Trans_To_Hill')
    trans_to_valley = pd.Series(dtype=float, name='Trans_To_Valley')
    trans_to_jasper = pd.Series(dtype=float, name='Trans_To_Jasper')
    
    if source_col and df_transfers['_Source_Norm'].notna().any():
        for loc in ['Hill', 'Valley', 'Jasper']:
            mask = df_transfers['_Source_Norm'] == loc
            if mask.any():
                trans_from = df_transfers[mask].groupby(sku_col)[qty_col].sum()
                if loc == 'Hill': trans_from_hill = trans_from
                elif loc == 'Valley': trans_from_valley = trans_from
                else: trans_from_jasper = trans_from
                log_func(f"  ✓ Transfers FROM {loc}: {len(df_transfers[mask])} records")
    
    if dest_col and df_transfers['_Dest_Norm'].notna().any():
        for loc in ['Hill', 'Valley', 'Jasper']:
            mask = df_transfers['_Dest_Norm'] == loc
            if mask.any():
                trans_to = df_transfers[mask].groupby(sku_col)[qty_col].sum()
                if loc == 'Hill': trans_to_hill = trans_to
                elif loc == 'Valley': trans_to_valley = trans_to
                else: trans_to_jasper = trans_to
                log_func(f"  ✓ Transfers TO {loc}: {len(df_transfers[mask])} records")
    
    return {
        'Trans_Hill': trans_from_hill,
        'Trans_Valley': trans_from_valley,
        'Trans_Jasper': trans_from_jasper,
        'Trans_To_Hill': trans_to_hill,
        'Trans_To_Valley': trans_to_valley,
        'Trans_To_Jasper': trans_to_jasper
    }


def _process_po_data(df_po, col_map):
    """Clean and summarize Purchase Order data."""
    if df_po.empty:
        return pd.Series(dtype=float, name='PO_Qty')
    
    po_sku_c = None
    for col in df_po.columns:
        if col_map['sku'].upper() in str(col).upper() or 'SKU' in str(col).upper():
            po_sku_c = col
            break
    if not po_sku_c:
        po_sku_c = df_po.columns[0]
    
    if po_sku_c in df_po.columns and 'Quantity ordered' in df_po.columns:
        df_po[po_sku_c] = df_po[po_sku_c].astype(str).str.replace(r'\.0$', '', regex=True)
        df_po['Quantity ordered'] = df_po['Quantity ordered'].apply(clean_currency)
        return df_po.groupby(po_sku_c)['Quantity ordered'].sum().rename('PO_Qty')
    
    return pd.Series(dtype=float, name='PO_Qty')


def find_header_row(file_path, search_terms=['AGLC SKU', 'SKU', 'Product']):
    """Scan first 20 rows to find header. For AGLC forms, header is typically at row 10 (0-indexed = 9)."""
    try:
        df_temp = pd.read_excel(file_path, header=None, nrows=20, engine='openpyxl')
        
        # First, check row 9 (which is row 10 in Excel - 0-indexed)
        if len(df_temp) > 9:
            row_9 = df_temp.iloc[9].astype(str).str.upper().tolist()
            # Check if this row contains SKU header
            if any('SKU' in str(cell).upper() for cell in row_9):
                return 9  # Row 10 in Excel (0-indexed = 9)
        
        # Fallback: scan for header
        for idx, row in df_temp.iterrows():
            row_str = row.astype(str).str.upper().tolist()
            # Skip rows that look like totals or summaries
            if any('TOTAL' in str(cell).upper() or 'SUMMARY' in str(cell).upper() for cell in row_str):
                continue
            for term in search_terms:
                if any(term.upper() in str(cell).upper() for cell in row_str):
                    return idx
    except Exception as e:
        pass
    return 9  # Default to row 10 (0-indexed = 9) for AGLC forms


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
                           Only inventory and sales are required. Others are optional.
        settings (dict): Configuration including logic rules and column mappings
        report_days (str): Number of days of sales data to analyze
        log_func (callable): Callback for logging messages
        finished_callback (callable): Callback when complete (bool success parameter)
    """
    try:
        log_func("--- Starting Analysis (v8.0) ---")
        
        rules_cannabis = settings.get('cannabis_logic', DEFAULT_SETTINGS['cannabis_logic'])
        rules_accessory = settings.get('accessory_logic', DEFAULT_SETTINGS['accessory_logic'])
        col_map = settings.get('column_mapping', DEFAULT_SETTINGS['column_mapping'])
        
        # Parse report days with validation
        try:
            days = float(report_days) if report_days else 30.0
            if days <= 0:
                days = 30.0
        except (ValueError, TypeError):
            days = 30.0
        
        weeks_factor = days / 7.0
        log_func(f"Analysis period: {days} days ({weeks_factor:.1f} weeks)")
        log_func("Loading Data Files...")
        
        # === 1. READ REQUIRED FILES ===
        if not file_paths.get('inventory'):
            raise ValueError("Inventory file is required!")
        if not file_paths.get('sales'):
            raise ValueError("Sales file is required!")
        
        df_inv = pd.read_csv(file_paths['inventory'])
        log_func(f"  ✓ Inventory: {len(df_inv)} records")
        
        df_sales = pd.read_csv(file_paths['sales'])
        log_func(f"  ✓ Sales: {len(df_sales)} records")
        
        # === 2. READ OPTIONAL FILES ===
        def load_opt(path, default_cols):
            """Load optional file or return empty DataFrame."""
            if path and os.path.exists(path):
                try:
                    return pd.read_csv(path)
                except Exception as e:
                    log_func(f"  ⚠️ Error loading {path}: {e}")
                    return pd.DataFrame(columns=default_cols)
            return pd.DataFrame(columns=default_cols)
        
        df_po = load_opt(file_paths.get('po'), [col_map['sku'], 'Quantity ordered'])
        if not df_po.empty:
            log_func(f"  ✓ Purchase Order: {len(df_po)} records")
        
        # Load all transfer files (can be multiple)
        transfer_files = []
        for loc_key in ['hill', 'valley', 'jasper']:
            if file_paths.get(loc_key):
                transfer_files.append(file_paths.get(loc_key))
        
        # Combine all transfer files and process
        df_transfers_list = []
        if transfer_files:
            for tf_path in transfer_files:
                if os.path.exists(tf_path):
                    try:
                        df_tf = pd.read_csv(tf_path)
                        if not df_tf.empty:
                            df_transfers_list.append(df_tf)
                            log_func(f"  ✓ Transfer file loaded: {os.path.basename(tf_path)} ({len(df_tf)} records)")
                    except Exception as e:
                        log_func(f"  ⚠️ Error loading transfer file {os.path.basename(tf_path)}: {e}")
        
        df_transfers_all = pd.concat(df_transfers_list, ignore_index=True) if df_transfers_list else pd.DataFrame()
        
        # AGLC manual form (optional - for case size/cost reference)
        df_aglc = pd.DataFrame(columns=['SKU', 'Case_Size', 'Case_Cost', 'New_SKU_This_Week', 'Available_Cases'])
        if file_paths.get('aglc') and os.path.exists(file_paths['aglc']):
            try:
                log_func("  -> Reading AGLC Manual Order Form...")
                header_row = find_header_row(file_paths['aglc'])
                log_func(f"  -> Using header row: {header_row + 1} (Excel row number)")
                df_aglc_raw = pd.read_excel(file_paths['aglc'], header=header_row, engine='openpyxl')
                
                # Try to find AGLC SKU column (flexible)
                sku_col = None
                for col in df_aglc_raw.columns:
                    if 'SKU' in str(col).upper():
                        sku_col = col
                        break
                if not sku_col:
                    sku_col = df_aglc_raw.columns[0]
                
                df_aglc_raw['SKU'] = df_aglc_raw[sku_col].astype(str).str.replace(r'\.0$', '', regex=True)
                
                # Try to find case size and cost columns
                size_col = None
                cost_col = None
                
                for col in df_aglc_raw.columns:
                    col_str = str(col).upper().strip()
                    if not size_col:
                        if ('EACH' in col_str and ('PER' in col_str or 'CASE' in col_str)) or \
                           ('UNITS' in col_str and 'CASE' in col_str) or \
                           ('CASE' in col_str and ('SIZE' in col_str or 'QTY' in col_str)):
                            size_col = col
                    if not cost_col:
                        if ('PRICE' in col_str and 'CASE' in col_str) or \
                           ('COST' in col_str and 'CASE' in col_str):
                            cost_col = col
                
                if size_col:
                    df_aglc_raw['Case_Size'] = pd.to_numeric(df_aglc_raw[size_col], errors='coerce').fillna(1)
                else:
                    df_aglc_raw['Case_Size'] = 1
                
                if cost_col:
                    df_aglc_raw['Case_Cost'] = pd.to_numeric(df_aglc_raw[cost_col].apply(clean_currency), errors='coerce').fillna(0)
                else:
                    df_aglc_raw['Case_Cost'] = 0
                
                # Find descriptive columns
                product_name_col = None
                category_col = None
                brand_col = None
                new_sku_col = None
                avail_cases_col = None
                
                for col in df_aglc_raw.columns:
                    col_str = str(col).upper().strip()
                    if not product_name_col and ('DESCRIPTION' in col_str or 'PRODUCT' in col_str): product_name_col = col
                    if not category_col and ('FORMAT' in col_str or 'CATEGORY' in col_str): category_col = col
                    if not brand_col and 'BRAND' in col_str: brand_col = col
                    if not new_sku_col and ('NEW' in col_str and 'SKU' in col_str): new_sku_col = col
                    if not avail_cases_col and ('AVAILABLE' in col_str and 'CASE' in col_str): avail_cases_col = col
                
                aglc_cols = ['SKU', 'Case_Size', 'Case_Cost']
                
                if product_name_col:
                    df_aglc_raw['Product Name'] = df_aglc_raw[product_name_col].astype(str)
                    aglc_cols.append('Product Name')
                if category_col:
                    df_aglc_raw['Category'] = df_aglc_raw[category_col].astype(str)
                    aglc_cols.append('Category')
                if brand_col:
                    df_aglc_raw['Brand'] = df_aglc_raw[brand_col].astype(str)
                    aglc_cols.append('Brand')
                if new_sku_col:
                    df_aglc_raw['New_SKU_This_Week'] = df_aglc_raw[new_sku_col]
                    aglc_cols.append('New_SKU_This_Week')
                if avail_cases_col:
                    df_aglc_raw['Available_Cases'] = pd.to_numeric(df_aglc_raw[avail_cases_col], errors='coerce').fillna(0)
                    aglc_cols.append('Available_Cases')
                
                df_aglc = df_aglc_raw[aglc_cols].drop_duplicates(subset=['SKU'])
                log_func(f"  ✓ AGLC Manual: {len(df_aglc)} products")
            except Exception as e:
                log_func(f"  ⚠️ AGLC file error: {e}")
        
        log_func("Processing Sales Data...")
        
        # === 3. CLEAN & PIVOT SALES ===
        cols_to_clean = [col_map['qty_sold'], col_map['profit'], col_map['net_sales'], col_map['gross_sales']]
        for col in cols_to_clean:
            if col in df_sales.columns:
                df_sales[col] = df_sales[col].apply(clean_currency)
        
        def normalize_loc(loc):
            s_loc = str(loc)
            if 'Hill' in s_loc: return 'Hill'
            if 'Valley' in s_loc: return 'Valley'
            if 'Jasper' in s_loc: return 'Jasper'
            return 'Other'
        
        df_sales['Loc_Key'] = df_sales['Location'].apply(normalize_loc) if 'Location' in df_sales.columns else 'Other'
        
        # Ensure date column is datetime and find last sale dates
        date_col = None
        for col in df_sales.columns:
            if 'DATE' in str(col).upper():
                date_col = col
                break
        
        if date_col:
            df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
            
        agg_map = {
            col_map['qty_sold']: 'sum',
            col_map['gross_sales']: 'sum',
            col_map['net_sales']: 'sum',
            col_map['profit']: 'sum'
        }
        if date_col:
            agg_map[date_col] = 'max'

        pivot_sales = df_sales.pivot_table(
            index=col_map['sku'],
            columns='Loc_Key', 
            values=list(agg_map.keys()),
            aggfunc=agg_map
        ).fillna(0)
        pivot_sales.columns = [f"{c[0]}_{c[1]}" for c in pivot_sales.columns]
        
        log_func("Merging Data...")
        # === 4. MERGE DATASETS ===
        master = df_inv.copy()
        
        if col_map.get('inventory_sku', 'SKU') != 'SKU':
            master.rename(columns={col_map['inventory_sku']: 'SKU'}, inplace=True)

        rename_map = {}
        for old, new in [(col_map.get('description'), 'Product Name'), (col_map.get('category'), 'Category'), (col_map.get('brand'), 'Brand')]:
            if old and old != new and old in master.columns: rename_map[old] = new
            
        if rename_map: master.rename(columns=rename_map, inplace=True)
        master['SKU'] = master['SKU'].astype(str).str.replace(r'\.0$', '', regex=True)
        master = pd.merge(master, pivot_sales, left_on='SKU', right_index=True, how='left').fillna(0)
        
        # Add New SKUs from AGLC
        if not df_aglc.empty:
            master_skus = set(master['SKU'].unique())
            aglc_skus = set(df_aglc['SKU'].unique())
            new_skus = aglc_skus - master_skus
            
            if new_skus:
                log_func(f"  ✓ Found {len(new_skus)} new SKUs from AGLC")
                new_sku_df = df_aglc[df_aglc['SKU'].isin(new_skus)].copy()
                for col in master.columns:
                    if col not in new_sku_df.columns:
                        if master[col].dtype.kind in ['i', 'f']: new_sku_df[col] = 0.0
                        else: new_sku_df[col] = ""
                master = pd.concat([master, new_sku_df], ignore_index=True)
            
            aglc_cols_to_merge = ['SKU', 'Case_Size', 'Case_Cost', 'New_SKU_This_Week', 'Available_Cases', 'Product Name', 'Category', 'Brand']
            aglc_subset = df_aglc[[c for c in aglc_cols_to_merge if c in df_aglc.columns]].copy()
            master = pd.merge(master, aglc_subset, on='SKU', how='left', suffixes=('', '_aglc'))
            
            for col in aglc_subset.columns:
                if col != 'SKU' and f'{col}_aglc' in master.columns:
                    master[col] = master[f'{col}_aglc'].combine_first(master.get(col, ""))
                    master.drop(columns=[f'{col}_aglc'], inplace=True)
        
        master['Case_Size'] = master['Case_Size'].fillna(1).replace(0, 1)
        master['Case_Cost'] = master['Case_Cost'].fillna(0)
        
        # Process PO and Transfer data
        po_series = _process_po_data(df_po, col_map)
        transfer_data = _process_transfer_data(df_transfers_all, log_func)

        # Join transfer data
        master = master.set_index('SKU')
        
        # Convert Series to DataFrame to avoid join conflicts
        transfer_df = pd.DataFrame({
            'Trans_Hill': transfer_data['Trans_Hill'],
            'Trans_Valley': transfer_data['Trans_Valley'],
            'Trans_Jasper': transfer_data['Trans_Jasper'],
            'Trans_To_Hill': transfer_data['Trans_To_Hill'],
            'Trans_To_Valley': transfer_data['Trans_To_Valley'],
            'Trans_To_Jasper': transfer_data['Trans_To_Jasper'],
            'PO_Qty': po_series
        })
        
        master = master.join(transfer_df, how='left').fillna(0)
        master = master.reset_index()
        
        log_func("Running Algorithms (Vectorized)...")
        
        # === 6. CALCULATE METRICS PER LOCATION ===
        master['Is_Accessory'] = ~master['SKU'].astype(str).str.upper().str.startswith("CNB-")
        master['Target_WOS'] = np.where(
            master['Is_Accessory'],
            rules_accessory['target_wos'],
            rules_cannabis['target_wos']
        )
        
        # Get PO destination from settings
        po_destination = settings.get('po_destination', 'J').upper()  # Default to Jasper
        
        # Calculate net PO based on destination
        # PO goes to destination location first
        if po_destination == 'H':
            master['PO_Net_Hill'] = master['PO_Qty']
            master['PO_Net_Valley'] = 0
            master['PO_Net_Jasper'] = 0
        elif po_destination == 'V':
            master['PO_Net_Valley'] = master['PO_Qty']
            master['PO_Net_Hill'] = 0
            master['PO_Net_Jasper'] = 0
        else:  # Default to Jasper
            master['PO_Net_Jasper'] = master['PO_Qty']
            master['PO_Net_Hill'] = 0
            master['PO_Net_Valley'] = 0
        
        log_func(f"  ✓ PO Destination: {po_destination} ({'Hill' if po_destination == 'H' else 'Valley' if po_destination == 'V' else 'Jasper'})")
        
        # Log transfer quantities
        if 'Trans_Hill' in master.columns and (master['Trans_Hill'] > 0).any():
            log_func(f"  ✓ Transfers FROM Hill: {master[master['Trans_Hill'] > 0]['Trans_Hill'].sum():.0f} units")
        if 'Trans_Valley' in master.columns and (master['Trans_Valley'] > 0).any():
            log_func(f"  ✓ Transfers FROM Valley: {master[master['Trans_Valley'] > 0]['Trans_Valley'].sum():.0f} units")
        if 'Trans_Jasper' in master.columns and (master['Trans_Jasper'] > 0).any():
            log_func(f"  ✓ Transfers FROM Jasper: {master[master['Trans_Jasper'] > 0]['Trans_Jasper'].sum():.0f} units")
        if 'Trans_To_Hill' in master.columns and (master['Trans_To_Hill'] > 0).any():
            log_func(f"  ✓ Transfers TO Hill: {master[master['Trans_To_Hill'] > 0]['Trans_To_Hill'].sum():.0f} units")
        if 'Trans_To_Valley' in master.columns and (master['Trans_To_Valley'] > 0).any():
            log_func(f"  ✓ Transfers TO Valley: {master[master['Trans_To_Valley'] > 0]['Trans_To_Valley'].sum():.0f} units")
        if 'Trans_To_Jasper' in master.columns and (master['Trans_To_Jasper'] > 0).any():
            log_func(f"  ✓ Transfers TO Jasper: {master[master['Trans_To_Jasper'] > 0]['Trans_To_Jasper'].sum():.0f} units")
        
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
                ('Sales' in c or 'Storage' in c or 'Inventory' in c or 'Qty' in c or 'Stock' in c) and
                c not in [col_sold, col_rev, col_gross, col_prof]
            ]
            
            for sc in stock_cols:
                if master[sc].dtype == object:
                    master[sc] = master[sc].apply(clean_currency)
            
            master[f'{loc}_Stock'] = master[stock_cols].sum(axis=1) if stock_cols else 0
            
            # Set incoming quantities based on PO destination and transfers
            # PO goes to destination location first
            # Transfers FROM a location reduce what stays at that location
            # Transfers FROM other locations TO this location increase incoming
            # For now, assume transfers FROM other locations go to the PO destination
            # (This can be refined if transfer files specify destination)
            
            if loc == 'Hill':
                # Hill receives: PO (if PO destination is Hill) + Transfers TO Hill
                incoming_po = master.get('PO_Net_Hill', 0)
                incoming_transfers = master.get('Trans_To_Hill', 0)
                # Transfers FROM Hill reduce what stays
                outgoing_transfers = master.get('Trans_Hill', 0)
                incoming_net = (incoming_po + incoming_transfers - outgoing_transfers).clip(lower=0)
                master[f'{loc}_Inc_Num'] = incoming_net
                
                # Build incoming string
                def build_inc_str(row):
                    parts = []
                    po_val = row.get('PO_Net_Hill', 0)
                    if po_val > 0:
                        parts.append(f"{int(po_val)} 📦")
                    trans_to = row.get('Trans_To_Hill', 0)
                    if trans_to > 0:
                        parts.append(f"{int(trans_to)} 🚚")
                    trans_from = row.get('Trans_Hill', 0)
                    if trans_from > 0:
                        inc_net = max(po_val + trans_to - trans_from, 0)
                        if inc_net == 0:
                            parts.append("(transferred out)")
                    return " + ".join(parts) if parts else "-"
                
                master[f'{loc}_Inc_Str'] = master.apply(build_inc_str, axis=1)
                
            elif loc == 'Valley':
                incoming_po = master.get('PO_Net_Valley', 0)
                incoming_transfers = master.get('Trans_To_Valley', 0)
                outgoing_transfers = master.get('Trans_Valley', 0)
                incoming_net = (incoming_po + incoming_transfers - outgoing_transfers).clip(lower=0)
                master[f'{loc}_Inc_Num'] = incoming_net
                
                def build_inc_str(row):
                    parts = []
                    po_val = row.get('PO_Net_Valley', 0)
                    if po_val > 0:
                        parts.append(f"{int(po_val)} 📦")
                    trans_to = row.get('Trans_To_Valley', 0)
                    if trans_to > 0:
                        parts.append(f"{int(trans_to)} 🚚")
                    trans_from = row.get('Trans_Valley', 0)
                    if trans_from > 0:
                        inc_net = max(po_val + trans_to - trans_from, 0)
                        if inc_net == 0:
                            parts.append("(transferred out)")
                    return " + ".join(parts) if parts else "-"
                
                master[f'{loc}_Inc_Str'] = master.apply(build_inc_str, axis=1)
                
            elif loc == 'Jasper':
                incoming_po = master.get('PO_Net_Jasper', 0)
                incoming_transfers = master.get('Trans_To_Jasper', 0)
                outgoing_transfers = master.get('Trans_Jasper', 0)
                incoming_net = (incoming_po + incoming_transfers - outgoing_transfers).clip(lower=0)
                master[f'{loc}_Inc_Num'] = incoming_net
                
                def build_inc_str(row):
                    parts = []
                    po_val = row.get('PO_Net_Jasper', 0)
                    if po_val > 0:
                        parts.append(f"{int(po_val)} 📦")
                    trans_to = row.get('Trans_To_Jasper', 0)
                    if trans_to > 0:
                        parts.append(f"{int(trans_to)} 🚚")
                    trans_from = row.get('Trans_Jasper', 0)
                    if trans_from > 0:
                        inc_net = max(po_val + trans_to - trans_from, 0)
                        if inc_net == 0:
                            parts.append("(transferred out)")
                    return " + ".join(parts) if parts else "-"
                
                master[f'{loc}_Inc_Str'] = master.apply(build_inc_str, axis=1)
            
            # Calculate report start date for time-aware velocity
            report_start = datetime.now() - pd.Timedelta(days=days)
            
            # Copy financial data (store Net Sales too for margin calculation)
            master[f'{loc}_Sold'] = master[col_sold]
            master[f'{loc}_Gross'] = master[col_gross]
            master[f'{loc}_Net'] = master[col_rev]  # Net Sales for margin calculation
            master[f'{loc}_Profit'] = master[col_prof]
            
            # Determine SOQ and Status
            from business_rules import InventoryMetrics, rules_dict_to_status_rules, calculate_soq, StatusDeterminer
            
            def apply_soq_and_status(row, loc_prefix, rules_can, rules_acc, report_days, start_date):
                is_accessory = row['Is_Accessory']
                rules_dict = rules_accessory if is_accessory else rules_cannabis
                rules = rules_dict_to_status_rules(rules_dict, is_accessory=is_accessory)
                
                # Get last sale date if available
                last_sale = row.get(f'{date_col}_{loc_prefix}') if date_col else None
                if pd.isna(last_sale) or last_sale == 0:
                    last_sale = None
                
                metrics = InventoryMetrics(
                    stock=max(0, int(row.get(f'{loc_prefix}_Stock', 0))),
                    incoming=max(0, int(row.get(f'{loc_prefix}_Inc_Num', 0))),
                    is_accessory=is_accessory,
                    total_units_sold=max(0, float(row.get(f'{col_map["qty_sold"]}_{loc_prefix}', 0))),
                    report_days=report_days,
                    report_start_date=start_date,
                    last_sale_date=last_sale
                )
                
                # Calculate velocity using new time-aware logic
                adj_velocity = metrics.calculate_velocity()
                
                # Assuming 'Case_Size' is a column in master, default to 1 if not found or invalid
                safe_case_size = max(int(row.get('Case_Size', 1)), 1)
                
                soq = calculate_soq(metrics, rules, safe_case_size)
                status = StatusDeterminer.determine_status(metrics, rules)
                
                # Calculate WOS based on adjusted velocity
                adj_wos = DEFAULT_SILENCE_THRESHOLD
                if adj_velocity > 0:
                    adj_wos = metrics.stock / adj_velocity
                elif metrics.stock == 0:
                    adj_wos = 0.0
                
                return pd.Series({'SOQ': soq, 'Status': status, 'Vel': adj_velocity, 'WOS': adj_wos})

            soq_status_df = master.apply(
                lambda row: apply_soq_and_status(row, loc, rules_cannabis, rules_accessory, days, report_start),
                axis=1
            )
            master[f'{loc}_SOQ'] = soq_status_df['SOQ']
            master[f'{loc}_Status'] = soq_status_df['Status']
            master[f'{loc}_Vel'] = soq_status_df['Vel']
            master[f'{loc}_WOS'] = soq_status_df['WOS']
            
            # Calculate margin percentage
            master[f'{loc}_Mrg'] = np.where(
                master[col_rev] > 0,
                master[col_prof] / master[col_rev],
                0.0
            )
            
            # Build StockDisplay: "5 + 12 🚚" format
            def build_stock_display(row):
                stock = int(row.get(f'{loc}_Stock', 0))
                incoming = int(row.get(f'{loc}_Inc_Num', 0))
                if incoming > 0:
                    return f"{stock} + {incoming} 🚚"
                elif stock > 0:
                    return str(stock)
                else:
                    return "0"
            master[f'{loc}_StockDisplay'] = master.apply(build_stock_display, axis=1)
            
            # Copy financial data
            master[f'{loc}_Sold'] = master[col_sold]
        
        log_func(f"✓ Calculated metrics for {len(master)} products")
        
        # === 7. WRITE EXCEL REPORT ===
        write_excel_report(master, rules_cannabis, rules_accessory, days, log_func)
        finished_callback(True)

    except Exception as e:
        log_func(f"❌ FATAL ERROR: {e}")
        import traceback
        error_trace = traceback.format_exc()
        log_func(f"Traceback:\n{error_trace}")
        traceback.print_exc()
        finished_callback(False)
