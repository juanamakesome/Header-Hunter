"""
Header Hunter v5.0 - History Builder Module
Builds and maintains a consolidated sales history from all historical CSV exports
"""
import os
import pandas as pd
import re
from pathlib import Path
from datetime import datetime


def extract_date_from_filename(filename):
    """
    Extract the end date from a product-sales filename.
    
    Expected format: product-sales-YYYY-MM-DD-YYYY-MM-DD.csv
    Returns the second date (the end date of the 30-day period).
    
    Args:
        filename (str): CSV filename
        
    Returns:
        datetime.date or None: End date of the report period
        
    Example:
        "product-sales-2025-11-01-2025-12-01.csv" → 2025-12-01
    """
    match = re.search(r'product-sales-(\d{4}-\d{2}-\d{2})-(\d{4}-\d{2}-\d{2})', filename)
    if match:
        end_date_str = match.group(2)  # Second date
        try:
            return pd.to_datetime(end_date_str).date()
        except ValueError:
            return None
    return None


def build_sales_history(history_folder, output_path, col_map, log_func=None):
    """
    Scan a folder of historical product-sales CSVs and build one consolidated history table.
    
    This function:
    1. Finds all product-sales-*.csv files in history_folder
    2. Extracts the report end date from each filename
    3. Reads relevant columns from each CSV
    4. Adds Report_End_Date column
    5. Concatenates all into one DataFrame
    6. Deduplicates (SKU + Report_End_Date pairs)
    7. Saves as parquet (faster, compressed) and CSV (readable)
    
    Args:
        history_folder (str): Path to folder containing historical product-sales CSVs
        output_path (str): Path where to save the history table (without extension)
        col_map (dict): Column mapping from settings (tells us what columns to expect)
        log_func (callable, optional): Function to log progress messages
        
    Returns:
        pd.DataFrame: The consolidated history table (also saved to disk)
    """
    if log_func is None:
        log_func = print
    
    history_folder = Path(history_folder)
    if not history_folder.exists():
        log_func(f"❌ History folder not found: {history_folder}")
        return None
    
    # Find all product-sales CSVs
    csv_files = sorted(history_folder.glob('product-sales-*.csv'))
    if not csv_files:
        log_func(f"⚠️ No product-sales CSVs found in {history_folder}")
        return None
    
    log_func(f"📚 Found {len(csv_files)} historical sales files")
    
    # Columns we need from each CSV
    sku_col = col_map.get('sku', 'SKU')
    qty_col = col_map.get('qty_sold', 'Quantity')
    net_sales_col = col_map.get('net_sales', 'Net sales')
    gross_sales_col = col_map.get('gross_sales', 'Gross sales')
    profit_col = col_map.get('profit', 'Profit')
    
    required_cols = [sku_col, qty_col, net_sales_col, gross_sales_col, profit_col]
    optional_cols = ['Location', 'Product Name', 'Category', 'Brand']
    
    frames = []
    
    for csv_path in csv_files:
        filename = csv_path.name
        report_end_date = extract_date_from_filename(filename)
        
        if report_end_date is None:
            log_func(f"⚠️ Could not parse date from {filename}, skipping")
            continue
        
        try:
            # Read CSV
            df = pd.read_csv(csv_path)
            
            # Check if required columns exist
            cols_to_use = []
            for col in required_cols:
                if col in df.columns:
                    cols_to_use.append(col)
            
            for col in optional_cols:
                if col in df.columns:
                    cols_to_use.append(col)
            
            if not cols_to_use:
                log_func(f"⚠️ {filename}: Could not find any recognized columns, skipping")
                continue
            
            # Select columns
            df_subset = df[cols_to_use].copy()
            
            # Add report date
            df_subset['Report_End_Date'] = report_end_date
            
            # Standardize column names to a consistent format
            rename_map = {
                sku_col: 'SKU',
                qty_col: 'Quantity',
                net_sales_col: 'Net_sales',
                gross_sales_col: 'Gross_sales',
                profit_col: 'Profit'
            }
            df_subset.rename(columns=rename_map, inplace=True)
            
            frames.append(df_subset)
            log_func(f"✅ {filename}: {len(df_subset)} rows")
            
        except Exception as e:
            log_func(f"❌ Error reading {filename}: {e}")
            continue
    
    if not frames:
        log_func("❌ No data was loaded from any CSV files")
        return None
    
    # Concatenate all
    log_func("🔄 Merging all files...")
    df_history = pd.concat(frames, ignore_index=True)
    
    # Deduplicate: keep latest version of each (SKU, Report_End_Date) pair
    log_func("🧹 Deduplicating...")
    df_history = df_history.drop_duplicates(
        subset=['SKU', 'Report_End_Date'],
        keep='last'
    )
    
    # Sort for easier inspection
    df_history = df_history.sort_values(['SKU', 'Report_End_Date']).reset_index(drop=True)
    
    log_func(f"📊 Final consolidated table: {len(df_history)} rows, {df_history['SKU'].nunique()} unique SKUs")
    
    # Save
    parquet_path = f"{output_path}.parquet"
    csv_path_out = f"{output_path}.csv"
    
    try:
        df_history.to_parquet(parquet_path, index=False)
        df_history.to_csv(csv_path_out, index=False)
        log_func(f"💾 Saved to {parquet_path} and {csv_path_out}")
    except Exception as e:
        log_func(f"❌ Error saving history: {e}")
        return None
    
    return df_history


def compute_rolling_velocities(df_history, sku, location=None, as_of_date=None, weeks_back=None):
    """
    Compute velocity metrics for a single SKU across different time windows.
    
    Given a history table and a SKU, this computes:
    - 4-week velocity
    - 12-week velocity
    - Lifetime velocity
    - Trend (comparison of recent vs older)
    
    Args:
        df_history (pd.DataFrame): Consolidated history table from build_sales_history()
        sku (str): SKU to analyze
        location (str, optional): Filter by location (if 'Location' column exists)
        as_of_date (datetime.date, optional): Latest date to consider (default: today)
        weeks_back (int, optional): How many weeks of data to examine (default: all)
        
    Returns:
        dict: Metrics including:
            - qty_4w, qty_12w, qty_lifetime
            - vel_4w, vel_12w, vel_lifetime (in units/week)
            - trend (recent vs older comparison)
    """
    if as_of_date is None:
        as_of_date = pd.Timestamp.now().date()
    
    # Filter to this SKU
    df_sku = df_history[df_history['SKU'] == sku].copy()
    
    if location and 'Location' in df_sku.columns:
        df_sku = df_sku[df_sku['Location'] == location]
    
    if len(df_sku) == 0:
        return {
            'qty_4w': 0, 'qty_12w': 0, 'qty_lifetime': 0,
            'vel_4w': 0, 'vel_12w': 0, 'vel_lifetime': 0,
            'trend': 'No data'
        }
    
    # Convert to datetime for filtering
    df_sku['Report_End_Date'] = pd.to_datetime(df_sku['Report_End_Date'])
    as_of_ts = pd.Timestamp(as_of_date)
    
    # 4-week window: last 28 days
    df_4w = df_sku[df_sku['Report_End_Date'] >= (as_of_ts - pd.Timedelta(days=28))]
    
    # 12-week window: last 84 days
    df_12w = df_sku[df_sku['Report_End_Date'] >= (as_of_ts - pd.Timedelta(days=84))]
    
    # Lifetime
    df_lifetime = df_sku
    
    # Compute quantities
    qty_4w = df_4w['Quantity'].sum()
    qty_12w = df_12w['Quantity'].sum()
    qty_lifetime = df_lifetime['Quantity'].sum()
    
    # Convert to velocities (units per week)
    vel_4w = qty_4w / 4.0 if len(df_4w) > 0 else 0
    vel_12w = qty_12w / 12.0 if len(df_12w) > 0 else 0
    vel_lifetime = qty_lifetime / (len(df_lifetime) * (7.0 / 7.0)) if len(df_lifetime) > 0 else 0
    
    # Trend: compare recent half vs older half of 12-week window
    trend = 'Stable'
    if len(df_12w) >= 2:
        mid_point = df_12w['Report_End_Date'].median()
        recent = df_12w[df_12w['Report_End_Date'] >= mid_point]['Quantity'].sum()
        older = df_12w[df_12w['Report_End_Date'] < mid_point]['Quantity'].sum()
        if older > 0:
            pct_change = ((recent - older) / older) * 100
            if pct_change > 20:
                trend = f'Growing (+{pct_change:.0f}%)'
            elif pct_change < -20:
                trend = f'Declining ({pct_change:.0f}%)'
    
    return {
        'qty_4w': qty_4w,
        'qty_12w': qty_12w,
        'qty_lifetime': qty_lifetime,
        'vel_4w': vel_4w,
        'vel_12w': vel_12w,
        'vel_lifetime': vel_lifetime,
        'trend': trend
    }


def load_or_build_history(history_folder, output_path, col_map, log_func=None, force_rebuild=False):
    """
    Load existing history table or build it if it doesn't exist.
    
    Checks for existing parquet file first (faster), falls back to CSV.
    If neither exists and history_folder has CSVs, builds the history.
    
    Args:
        history_folder (str): Path to folder with historical CSVs
        output_path (str): Where to save/load consolidated history
        col_map (dict): Column mapping from settings
        log_func (callable, optional): Logging function
        force_rebuild (bool): If True, always rebuild from CSVs (default: False)
        
    Returns:
        pd.DataFrame or None: Consolidated history table
    """
    if log_func is None:
        log_func = print
    
    parquet_path = f"{output_path}.parquet"
    csv_path = f"{output_path}.csv"
    
    # Try to load existing
    if not force_rebuild:
        if os.path.exists(parquet_path):
            try:
                log_func(f"📖 Loading existing history from {parquet_path}")
                return pd.read_parquet(parquet_path)
            except Exception as e:
                log_func(f"⚠️ Error loading parquet: {e}, trying CSV")
        
        if os.path.exists(csv_path):
            try:
                log_func(f"📖 Loading existing history from {csv_path}")
                return pd.read_csv(csv_path)
            except Exception as e:
                log_func(f"⚠️ Error loading CSV: {e}")
    
    # Build from scratch
    log_func("🔨 Building history from CSVs...")
    return build_sales_history(history_folder, output_path, col_map, log_func)