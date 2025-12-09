"""
Header Hunter v5.0 - Utilities Module
Configuration management and resource path handling
"""
import sys
import os
import json

APP_TITLE = "Header Hunter v5.0 (Modular)"
CONFIG_FILE = 'header_hunter_config.json'

# Constant threshold values with clear meaning
DEFAULT_SILENCE_THRESHOLD = 999.0  # WOS value when no velocity data available

DEFAULT_SETTINGS = {
    "cannabis_logic": {
        "hot_velocity": 2.0,           # Units/week threshold for "Hot" status
        "reorder_point": 2.5,          # Minimum weeks-on-stock before reorder
        "target_wos": 4.0,             # Target weeks of stock to maintain
        "dead_wos": 26,                # Weeks threshold for "Dead" classification
        "dead_on_hand": 5              # Minimum on-hand units for "Dead" status
    },
    "accessory_logic": {
        "hot_velocity": 0.5,           # Lower threshold for accessories
        "reorder_point": 4.0,
        "target_wos": 8.0,             # Accessories held longer
        "dead_wos": 52,                # Double the cannabis threshold
        "dead_on_hand": 3
    },
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