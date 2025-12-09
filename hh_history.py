"""
Header Hunter - History Module
"The Historian"
Reads the Immaculate Memory Bank (Parquet) created by Ingest.
"""
import pandas as pd
from pathlib import Path

# MATCH THIS NAME EXACTLY TO hh_ingest.py
MASTER_DB_NAME = "sales_history_master.parquet"

def load_history_db(history_folder, log_func=None):
    """
    Loads the master parquet file directly.
    """
    if log_func is None: 
        log_func = print
        
    history_folder = Path(history_folder)
    parquet_path = history_folder / MASTER_DB_NAME
    
    if not parquet_path.exists():
        log_func(f"⚠️ Memory Bank Empty! Could not find: {parquet_path}")
        log_func("   -> Please run 'hh_ingest.py' to feed the bank.")
        return None
        
    try:
        log_func(f"📚 Opening Memory Bank: {MASTER_DB_NAME}")
        df = pd.read_parquet(parquet_path)
        log_func(f"   -> Loaded {len(df)} historical records.")
        return df
    except Exception as e:
        log_func(f"❌ Error reading Memory Bank: {e}")
        return None

def compute_rolling_velocities(df_history, sku, location=None, as_of_date=None):
    """
    Standard Logic: Looks for the most relevant snapshot in the history.
    """
    if as_of_date is None:
        as_of_date = pd.Timestamp.now()
    else:
        as_of_date = pd.Timestamp(as_of_date)
    
    # Filter for SKU and Location
    mask = (df_history['SKU'] == sku)
    if location:
        mask &= (df_history['Location'] == location)
    
    df_sku = df_history[mask].copy()
    
    if df_sku.empty:
        return {'vel_4w': 0, 'vel_12w': 0, 'trend': 'No Data'}
    
    # Sort by date descending (newest first)
    df_sku = df_sku.sort_values('Report_End_Date', ascending=False)
    
    # --- HELPER: Get velocity from a specific point in time ---
    def get_velocity_at_date(target_date, tolerance_days=45):
        # Calculate time difference
        time_diffs = (df_sku['Report_End_Date'] - target_date).abs()
        
        # Get the closest report within tolerance
        valid_mask = time_diffs <= pd.Timedelta(days=tolerance_days)
        valid_reports = df_sku[valid_mask]
        
        if valid_reports.empty:
            return 0.0, 0.0
        
        # Find the index of the closest report
        closest_idx = time_diffs[valid_mask].idxmin()
        best_row = df_sku.loc[closest_idx]
        
        qty = best_row['Quantity']
        
        # Default to 30 days if not specified, to calculate weekly run rate
        days_in_report = 30.0 
        velocity = (qty / days_in_report) * 7.0
        return qty, velocity
    
    # 1. Current Velocity (Closest to Today)
    qty_now, vel_now = get_velocity_at_date(as_of_date)
    
    # 2. Past Velocity (Closest to 12 weeks ago)
    date_12w_ago = as_of_date - pd.Timedelta(weeks=12)
    qty_old, vel_old = get_velocity_at_date(date_12w_ago)
    
    # 3. Calculate Trend
    trend = "Stable"
    if vel_old > 0:
        pct_change = ((vel_now - vel_old) / vel_old)
        if pct_change > 0.25:
            trend = f"Growing (+{pct_change:.0%})"
        elif pct_change < -0.25:
            trend = f"Declining ({pct_change:.0%})"
    elif vel_now > 0 and vel_old == 0:
        trend = "New / Spiking"
    
    return {
        'vel_4w': vel_now,
        'vel_12w': vel_old,
        'trend': trend
    }