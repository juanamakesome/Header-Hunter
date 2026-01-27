"""
Header Hunter v8.0 - Utilities Module
Configuration management and resource path handling
"""
import sys
import os
import json
from pathlib import Path
from typing import Optional, Dict, Any

APP_TITLE = "🎯 Header Hunter v8.0 | Cockpit"
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
        "description": "Product Name",   # Maps CSV header to standard name (e.g., "Description" -> "Product Name")
        "category": "Category",          # Maps CSV header to standard name (e.g., "Type" -> "Category")
        "brand": "Brand",                # Maps CSV header to standard name (e.g., "Producer" -> "Brand")
        "qty_sold": "Quantity",
        "net_sales": "Net sales",
        "gross_sales": "Gross sales",
        "profit": "Profit",
        "inventory_sku": "SKU"
    }
}


def resource_path(relative_path: str) -> str:
    """
    Get absolute path to resource file. Works for both development and PyInstaller bundled builds.
    
    Args:
        relative_path: Path relative to this module or bundle root
        
    Returns:
        Absolute path to the resource
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Fallback for development environment
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)


def resolve_data_path(config_path: str) -> Optional[str]:
    """
    Resolve configuration path to existing directory.
    
    This function attempts to find the data folder in order:
    1. Direct path (if it exists)
    2. Relative to application directory
    3. User's Documents/HeaderHunter folder
    
    Args:
        config_path: Path specified in config file or user input
        
    Returns:
        Valid path string if found, None otherwise
    """
    if not config_path:
        return None
    
    # Clean up path
    config_path = config_path.strip()
    
    # Try exact path first
    if os.path.exists(config_path):
        return config_path
    
    # Try relative to application directory
    app_dir = os.path.dirname(os.path.abspath(__file__))
    rel_path = os.path.join(app_dir, config_path)
    if os.path.exists(rel_path):
        return rel_path
    
    # Try relative to parent (if app is in subdirectory)
    parent_path = os.path.join(os.path.dirname(app_dir), config_path)
    if os.path.exists(parent_path):
        return parent_path
    
    # Try Documents folder
    try:
        docs_path = os.path.join(
            os.path.expanduser('~'), 
            'Documents', 
            'HeaderHunter',
            os.path.basename(config_path)
        )
        if os.path.exists(docs_path):
            return docs_path
    except Exception:
        pass
    
    return None


def validate_file_paths(paths: Dict[str, str]) -> Dict[str, bool]:
    """
    Validate that all configured file paths exist.
    
    Args:
        paths: Dictionary of file paths from config
        
    Returns:
        Dictionary mapping file keys to existence status
    """
    status = {}
    for key, path in paths.items():
        if path:
            status[key] = os.path.exists(path)
        else:
            status[key] = False
    return status


def load_config() -> Dict[str, Any]:
    """
    Load application configuration from JSON file.
    
    Checks in this order:
    1. Local directory (next to .exe) - for user persistence
    2. Bundled resources (sys._MEIPASS) - for default settings
    
    Returns:
        Configuration dictionary with 'settings' key
    """
    # 1. Try local config first
    config_to_load = CONFIG_FILE
    
    if not os.path.exists(config_to_load):
        # 2. Fallback to bundled config
        config_to_load = resource_path(CONFIG_FILE)
        
    if os.path.exists(config_to_load):
        try:
            with open(config_to_load, 'r') as f:
                data = json.load(f)
            
            # Ensure required structure exists
            if 'settings' not in data:
                data['settings'] = DEFAULT_SETTINGS.copy()
            
            # Clean up paths: remove ones that don't exist (e.g., Spudn paths)
            if 'paths' in data:
                valid_paths = {}
                for key, path in data['paths'].items():
                    if path and os.path.exists(path):
                        valid_paths[key] = path
                    # Skip invalid paths - user will select folder on first run
                data['paths'] = valid_paths
            else:
                data['paths'] = {}
            
            return data
            
        except (json.JSONDecodeError, IOError) as e:
            # Silently fall back to defaults on read error
            return {'settings': DEFAULT_SETTINGS.copy(), 'paths': {}}
    else:
        return {'settings': DEFAULT_SETTINGS.copy(), 'paths': {}}


def save_config(data: Dict[str, Any]) -> bool:
    """
    Save application configuration to JSON file.
    
    Args:
        data: Configuration dictionary to save
        
    Returns:
        True if saved successfully, False otherwise
    """
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(data, f, indent=4)
        return True
    except (IOError, OSError) as e:
        # Log but don't crash - user will re-select folder on next run
        logger.warning(f"Could not save config: {e}")
        return False


def create_empty_config() -> Dict[str, Any]:
    """
    Create a clean configuration with no paths (forces first-run setup).
    
    Returns:
        Configuration with default settings and empty paths
    """
    return {
        'settings': DEFAULT_SETTINGS.copy(),
        'paths': {}
    }