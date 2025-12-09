"""
Header Hunter- Utilities Module
Configuration management and resource path handling
"""
import sys
import os
import json
import re
import pandas as pd

APP_TITLE = "Header Hunter v4.20 (Kowalski Protocol)"
CONFIG_FILE = 'header_hunter_config.json'

# Constant threshold values with clear meaning
DEFAULT_SILENCE_THRESHOLD = 999.0  # WOS value when no velocity data available

DEFAULT_SETTINGS = {
    "cannabis_logic": {
        "hot_velocity": 2.0,           # Units/week threshold for "Hot" status
        "reorder_point": 2.5,          # Minimum weeks of stock before reorder
        "target_wos": 4.0,             # Target weeks of stock to maintain
        "dead_wos": 26,                # Weeks threshold for "Dead" classification
        "dead_on_hand": 5              # Minimum on-hand units for "Dead" status
    },
    "accessory_logic": {
        "hot_velocity": 0.5,           
        "reorder_point": 4.0,          
        "target_wos": 8.0,             
        "dead_wos": 52,                
        "dead_on_hand": 3
    },
    "history_folder": "",               # Optional folder for historical sales CSVs
    "column_mapping": {
        "sku": "SKU",
        "description": "Product Name",
        "qty_sold": "Quantity",
        "net_sales": "Net sales",
        "gross_sales": "Gross sales",
        "profit": "Profit",
        "inventory_sku": "SKU"
    }
}


def resource_path(relative_path):
    """
    Get absolute path to resource file. Works for both development and PyInstaller bundled builds.
    
    Args:
        relative_path (str): Path relative to this module or bundle root
        
    Returns:
        str: Absolute path to the resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Fallback for development environment
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)


def load_config():
    """
    Load application configuration from JSON file.
    Returns default settings if file doesn't exist or is corrupted.
    
    Returns:
        dict: Configuration dictionary with 'settings' key
    """
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                data = json.load(f)
                # Ensure required structure exists
                if 'settings' not in data:
                    data['settings'] = DEFAULT_SETTINGS
                return data
        except (json.JSONDecodeError, IOError) as e:
            # Silently fall back to defaults on read error
            return {'settings': DEFAULT_SETTINGS}
    else:
        return {'settings': DEFAULT_SETTINGS}


def save_config(data):
    """
    Save application configuration to JSON file.
    Silently fails if file cannot be written (e.g., permission error).
    
    Args:
        data (dict): Configuration dictionary to save
    """
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(data, f, indent=4)
    except (IOError, OSError) as e:
        # Silently ignore write failures to prevent UI crashes
        pass


def normalize_sku(sku_value):
    """
    Standard SKU normalization (use everywhere).
    - Converts to string
    - Removes trailing .0 from floats
    - Converts to uppercase
    - Strips whitespace
    - Returns None if invalid
    """
    if pd.isna(sku_value):
        return None
    
    normalized = str(sku_value).strip()
    
    # Remove float decimal artifact
    normalized = re.sub(r'\.0$', '', normalized)
    
    # Uppercase
    normalized = normalized.upper()
    
    # Validate: must have alphanumeric
    if not re.search(r'[A-Z0-9]', normalized):
        return None
    
    return normalized


def normalize_location(location_str):
    """
    Normalize location strings with case-insensitive matching.
    Returns standardized location name or 'Other'.
    Logs unmapped values for quality control.
    """
    if pd.isna(location_str):
        return 'Other'
    
    s = str(location_str).lower().strip()
    
    # Direct name matches (most specific first to avoid false positives)
    if 'jasper' in s:
        return 'Jasper'
    elif 'valley' in s:
        return 'Valley'
    elif 'hill' in s:
        return 'Hill'
    else:
        # Flag unmapped for review
        if s not in ['', 'other', 'unknown', 'n/a', '-']:
            return f'Other_UNMAPPED[{s[:20]}]'
        return 'Other'


def validate_column_mapping(df, col_map, expected_keys, log_func=None):
    """
    Validates that required columns exist in dataframe.
    
    Args:
        df: pandas DataFrame to check
        col_map: dict mapping logical names to actual column names
        expected_keys: list of keys that MUST be in col_map
        log_func: logging function (default: print)
    
    Returns:
        (valid: bool, missing: list, rename_map: dict)
    """
    if log_func is None:
        log_func = print
    
    missing = []
    rename_map = {}
    
    for key in expected_keys:
        source_col = col_map.get(key)
        
        if not source_col:
            missing.append(f"No mapping for '{key}'")
            continue
        
        if source_col not in df.columns:
            missing.append(f"'{key}' mapped to '{source_col}' but column not found")
        else:
            rename_map[source_col] = key.upper()
    
    if missing:
        log_func(f"⚠️  Column Mapping Issues:")
        for m in missing:
            log_func(f"   - {m}")
        log_func(f"   Available columns: {list(df.columns)}")
        return False, missing, rename_map
    
    return True, [], rename_map