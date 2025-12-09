"""
Header Hunter - Ingest Module
"The Librarian"
Now scouts the Downloads folder directly.
"""
import pandas as pd
import os
import shutil
import re
from datetime import datetime
from pathlib import Path
from hh_utils import load_config

MASTER_DB_NAME = "sales_history_master.parquet" 

def extract_date(filename):
    """Extracts the END DATE from the filename."""
    matches = re.findall(r'\d{4}-\d{2}-\d{2}', filename)
    if len(matches) >= 2:
        return pd.to_datetime(matches[1])
    return None

def ingest_file(filepath, log_func=print):
    """Reads a raw CSV, standardizes it, and prepares it for the bank."""
    config = load_config()
    col_map = config.get('settings', {}).get('column_mapping', {})
    
    filename = os.path.basename(filepath)
    report_date = extract_date(filename)
    
    if not report_date:
        log_func(f"⚠️ SKIPPING: Could not find valid date range in {filename}")
        return None

    log_func(f"📖 Reading Snapshot: {filename} (Date: {report_date.date()})")
    
    try:
        df = pd.read_csv(filepath)
        
        # Map Columns
        rename_map = {
            col_map.get('sku', 'SKU'): 'SKU',
            col_map.get('qty_sold', 'Quantity'): 'Quantity',
            col_map.get('net_sales', 'Net sales'): 'Net_sales',
            col_map.get('gross_sales', 'Gross sales'): 'Gross_sales',
            col_map.get('profit', 'Profit'): 'Profit'
        }
        actual_rename = {k: v for k, v in rename_map.items() if k in df.columns}
        df.rename(columns=actual_rename, inplace=True)
        
        # Normalize
        df['SKU'] = df['SKU'].astype(str).str.replace(r'\.0$', '', regex=True).str.upper()
        df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        
        # Location
        if 'Location' in df.columns:
            def clean_loc(x):
                s = str(x).lower()
                if 'hill' in s: return 'Hill'
                if 'valley' in s: return 'Valley'
                if 'jasper' in s: return 'Jasper'
                return 'Other'
            df['Location'] = df['Location'].apply(clean_loc)
        else:
            df['Location'] = 'Other'
            
        # Stamp Time
        df['Report_End_Date'] = report_date
        
        # Filter Columns
        keep_cols = ['SKU', 'Quantity', 'Location', 'Report_End_Date', 'Net_sales', 'Gross_sales', 'Profit']
        return df[[c for c in keep_cols if c in df.columns]]

    except Exception as e:
        log_func(f"❌ Error processing {filename}: {e}")
        return None

def update_memory_bank(log_func=print):
    """The main ritual."""
    # 1. Load Configuration
    config = load_config()
    history_root = config.get('settings', {}).get('history_folder', '')
    
    if not history_root or not os.path.exists(history_root):
        log_func("❌ Error: Memory Bank Folder is not set!")
        log_func("   -> Go to Settings and select your 'MEMORYBANKPLZNOTOUCH' folder.")
        return

    # 2. Determine Source Folder (The "Inbox")
    # First check if user set a custom 'ingest_folder' in settings
    source_dir = config.get('settings', {}).get('ingest_folder', '')
    
    # If not set, Default to standard Windows/Mac Downloads folder
    if not source_dir or not os.path.exists(source_dir):
        source_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
    
    log_func(f"🕵️ Scouting for reports in: {source_dir}")

    # Define Archive and Master Paths
    ARCHIVE_FOLDER = os.path.join(history_root, "Archive")
    MASTER_PATH = os.path.join(history_root, MASTER_DB_NAME)

    if not os.path.exists(ARCHIVE_FOLDER):
        os.makedirs(ARCHIVE_FOLDER)

    # --- 3. Load Master ---
    if os.path.exists(MASTER_PATH):
        try:
            master_df = pd.read_parquet(MASTER_PATH)
            # log_func(f"💾 Master Bank Loaded ({len(master_df)} records)")
        except Exception as e:
            log_func(f"⚠️ Master DB Corrupt ({e}). Starting Fresh.")
            master_df = pd.DataFrame()
    else:
        log_func("✨ Creating new Master Bank.")
        master_df = pd.DataFrame()

    # --- 4. Find Files in Source ---
    # We strictly look for 'product-sales' CSVs to avoid touching other downloads
    files = [f for f in os.listdir(source_dir) if f.startswith('product-sales') and f.endswith('.csv')]
    
    if not files:
        log_func(f"💤 No 'product-sales' CSVs found in Downloads.")
        return

    updates_made = False
    
    for f in files:
        full_path = os.path.join(source_dir, f)
        
        # SKIP if the file is currently open/locked (basic check)
        try:
            os.rename(full_path, full_path)
        except OSError:
            log_func(f"⚠️ Skipping {f} (File is open/locked)")
            continue

        new_data = ingest_file(full_path, log_func)
        
        if new_data is not None:
            # --- 5. SMART MERGE ---
            report_date = new_data['Report_End_Date'].iloc[0]
            
            if not master_df.empty:
                pre_count = len(master_df)
                master_df = master_df[master_df['Report_End_Date'] != report_date]
                removed = pre_count - len(master_df)
                if removed > 0:
                    log_func(f"   ♻️  Overwriting {removed} records for {report_date.date()}")
            
            master_df = pd.concat([master_df, new_data], ignore_index=True)
            updates_made = True
            
            # --- 6. Archive ---
            try:
                shutil.move(full_path, os.path.join(ARCHIVE_FOLDER, f))
                log_func(f"   📦 Moved to Archive: {f}")
            except Exception as e:
                log_func(f"   ⚠️ Could not move file: {e}")

    # --- 7. Save Master ---
    if updates_made:
        master_df = master_df.sort_values(by=['Report_End_Date', 'SKU'], ascending=[False, True])
        master_df.to_parquet(MASTER_PATH, index=False)
        log_func(f"✅ Memory Bank Saved! Total Records: {len(master_df)}")
    else:
        log_func("No valid updates processed.")

if __name__ == "__main__":
    update_memory_bank()
