"""
Header Hunter v8.0 - The Cockpit

Main entry point with dependency validation and error handling.
"""

import sys
from pathlib import Path
from typing import Tuple, List


def validate_dependencies() -> Tuple[bool, List[str]]:
    """
    Check if all required dependencies are installed.
    
    Returns:
        Tuple of (all_ok, missing_packages)
    """
    missing = []
    
    # Core dependencies
    try:
        import customtkinter  # noqa: F401
    except ImportError:
        missing.append("customtkinter")
    
    try:
        import pandas  # noqa: F401
    except ImportError:
        missing.append("pandas")
    
    try:
        import xlsxwriter  # noqa: F401
    except ImportError:
        missing.append("xlsxwriter")
    
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        missing.append("openpyxl")
    
    return len(missing) == 0, missing


def validate_local_modules() -> Tuple[bool, List[str]]:
    """
    Check if all local modules exist. Skip if running as bundled executable.
    
    Returns:
        Tuple of (all_ok, missing_modules)
    """
    import sys
    if getattr(sys, 'frozen', False):
        return True, []
        
    missing = []
    project_root = Path(__file__).resolve().parent
    
    required_modules = [
        'hh_gui_modern',
        'hh_logic',
        'hh_utils',
        'business_rules',
        'excel_column_map',
        'logging_config'
    ]
    
    for module_name in required_modules:
        module_path = project_root / f"{module_name}.py"
        if not module_path.exists():
            missing.append(f"{module_name}.py")
    
    return len(missing) == 0, missing


def main() -> bool:
    """Initialize and start the application."""
    
    try:
        # Step 1: Validate dependencies
        print("🔍 Checking dependencies...")
        deps_ok, missing_deps = validate_dependencies()
        
        if not deps_ok:
            print("❌ Missing required packages:")
            for pkg in missing_deps:
                print(f"   - {pkg}")
            print("\nFix with: pip install -r requirements.txt")
            return False
        
        print("✓ Core dependencies found")
        
        # Step 2: Validate local modules
        print("🔍 Checking local modules...")
        modules_ok, missing_modules = validate_local_modules()
        
        if not modules_ok:
            print("❌ Missing local modules:")
            for mod in missing_modules:
                print(f"   - {mod}")
            return False
        
        print("✓ Local modules found")
        
        # Step 3: Validate configuration
        print("🔍 Checking configuration...")
        try:
            from hh_utils import load_config, APP_TITLE
            config = load_config()
            if not config.get('settings'):
                print("⚠️ Configuration incomplete, using defaults")
            print("✓ Configuration loaded")
        except Exception as e:
            print(f"⚠️ Configuration warning: {e}")
            print("   Continuing with defaults...")
        
        # Step 4: Launch GUI
        print(f"🚀 Starting {APP_TITLE}...")
        from hh_gui_modern import HeaderHunterCockpit
        
        app = HeaderHunterCockpit()
        app.mainloop()
        
        return True
        
    except ImportError as e:
        print(f"❌ Import Error: {e}")
        print("\nTroubleshooting:")
        print("1. Ensure all required packages: pip install -r requirements.txt")
        print("2. Check that all .py files exist in the project directory")
        print("3. Verify Python version is 3.8 or higher: python --version")
        return False
        
    except Exception as e:  # noqa: BLE001
        print(f"❌ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        print("\nPlease report this error to the development team.")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
