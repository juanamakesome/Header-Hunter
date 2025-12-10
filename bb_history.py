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
    FIXED: Computes rolling velocities using ACTUAL report date ranges.
    
    Changes:
    - Groups by Report_End_Date first to avoid double-counting
    - Calculates actual day span between reports
    - Returns confidence score
    - Handles edge cases better
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
        return {
            'vel_4w': 0,
            'vel_12w': 0,
            'trend': 'No Data',
            'confidence': 0,
            'data_points': 0
        }
    
    # FIXED: Group by report date to avoid duplicate counting
    df_sku = df_sku.groupby('Report_End_Date').agg({
        'Quantity': 'sum'
    }).reset_index()
    
    df_sku = df_sku.sort_values('Report_End_Date', ascending=False)
    data_points = len(df_sku)
    
    def get_velocity_at_date(target_date, tolerance_days=45):
        """Get velocity nearest to target date."""
        time_diffs = (df_sku['Report_End_Date'] - target_date).abs()
        valid_mask = time_diffs <= pd.Timedelta(days=tolerance_days)
        valid_reports = df_sku[valid_mask]
        
        if valid_reports.empty:
            return 0.0, 0.0, 0
        
        # Get closest report
        closest_idx = time_diffs[valid_mask].idxmin()
        best_row = df_sku.loc[closest_idx]
        
        qty = best_row['Quantity']
        report_date = best_row['Report_End_Date']
        
        # FIXED: Calculate actual report span
        # Find the position of this report_date in the sorted dataframe
        position_int = None
        for idx in range(len(df_sku)):
            if df_sku.iloc[idx]['Report_End_Date'] == report_date:
                position_int = idx
                break
        
        if position_int is None:
            position_int = 0
        
        if position_int == 0:
            # Most recent report - assume 30 days
            days_in_report = 30.0
            confidence = 85
        elif position_int < len(df_sku) - 1:
            # Calculate days between reports
            # Get the next report (position + 1) - since sorted descending, this is the older report
            next_report = df_sku.iloc[position_int + 1]['Report_End_Date']
            days_in_report = (report_date - next_report).days
            confidence = 80 if 25 <= days_in_report <= 35 else 60
        else:
            # Oldest report - estimate 30 days
            days_in_report = 30.0
            confidence = 50
        
        # Ensure minimum 1 day to avoid division by zero
        days_in_report = max(days_in_report, 1)
        
        # Convert to weekly velocity
        velocity = (qty / days_in_report) * 7.0
        
        return qty, velocity, confidence
    
    # 1. Current Velocity (Closest to Today)
    qty_now, vel_now, conf_now = get_velocity_at_date(as_of_date)
    
    # 2. Past Velocity (Closest to 12 weeks ago)
    date_12w_ago = as_of_date - pd.Timedelta(weeks=12)
    qty_old, vel_old, conf_old = get_velocity_at_date(date_12w_ago)
    
    # 3. Calculate Trend
    trend = "Stable"
    if vel_old > 0:
        pct_change = ((vel_now - vel_old) / vel_old)
        if pct_change > 0.25:
            trend = f"↑ Growing (+{pct_change:.0%})"
        elif pct_change < -0.25:
            trend = f"↓ Declining ({pct_change:.0%})"
    elif vel_now > 0 and vel_old == 0:
        trend = "✨ New / Spiking"
    
    overall_confidence = min(conf_now, conf_old)
    
    return {
        'vel_4w': vel_now,
        'vel_12w': vel_old,
        'trend': trend,
        'confidence': overall_confidence,
        'data_points': data_points
    }