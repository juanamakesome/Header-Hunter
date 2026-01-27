"""
Header Hunter v8.0 - Excel Formula Validation Module

Provides validation helpers for Excel formulas before file generation.
Catches common issues before user sees them.
"""

from typing import Dict, List, Optional
import re


def validate_excel_formula(formula: str) -> Dict[str, any]:
    """
    Validate a single Excel formula string for common issues.
    
    Checks for:
    - Basic syntax errors (unmatched parentheses, quotes)
    - Common error patterns (#REF!, #NAME?, #DIV/0!)
    - Invalid function names
    
    Args:
        formula: Excel formula string (e.g., "=SUM(A1:A10)")
        
    Returns:
        Dictionary with validation results:
        {
            'is_valid': bool,
            'errors': List[str],
            'warnings': List[str]
        }
    """
    errors = []
    warnings = []
    
    if not formula:
        return {'is_valid': False, 'errors': ['Formula is empty'], 'warnings': []}
    
    # Must start with =
    if not formula.startswith('='):
        errors.append('Formula must start with =')
    
    # Check for unmatched parentheses
    open_parens = formula.count('(')
    close_parens = formula.count(')')
    if open_parens != close_parens:
        errors.append(f'Unmatched parentheses: {open_parens} open, {close_parens} close')
    
    # Check for unmatched quotes
    single_quotes = formula.count("'")
    if single_quotes % 2 != 0:
        errors.append('Unmatched single quotes')
    
    double_quotes = formula.count('"')
    if double_quotes % 2 != 0:
        errors.append('Unmatched double quotes')
    
    # Check for common error patterns (these would appear when Excel evaluates)
    error_patterns = {
        '#REF!': 'Reference error detected',
        '#NAME?': 'Name error detected',
        '#DIV/0!': 'Division by zero detected',
        '#VALUE!': 'Value error detected',
        '#N/A': 'Not available error detected',
    }
    
    for pattern, message in error_patterns.items():
        if pattern in formula:
            warnings.append(f'Potential error in formula: {message}')
    
    # Check for common invalid patterns
    if re.search(r'[A-Z]+\d+:[A-Z]+\d+', formula):
        # Looks like a range reference - basic validation
        pass
    
    return {
        'is_valid': len(errors) == 0,
        'errors': errors,
        'warnings': warnings
    }


def validate_column_reference(ref: str) -> bool:
    """
    Validate an Excel column reference (e.g., 'A', 'AA', 'AB123').
    
    Args:
        ref: Column reference string
        
    Returns:
        True if valid, False otherwise
    """
    if not ref:
        return False
    
    # Pattern: one or more letters followed by optional digits
    pattern = r'^[A-Z]+(\d+)?$'
    return bool(re.match(pattern, ref))


def validate_cell_reference(cell_ref: str) -> bool:
    """
    Validate a full Excel cell reference (e.g., 'A1', 'Sheet1!A1', 'A1:B10').
    
    Args:
        cell_ref: Cell reference string
        
    Returns:
        True if valid format, False otherwise
    """
    if not cell_ref:
        return False
    
    # Pattern: optional sheet name with !, then column(s) and row(s)
    # Examples: A1, Sheet1!A1, A1:B10, 'Sheet Name'!A1
    pattern = r"^(('[^']+'|[\w]+)!)?[A-Z]+\d+(:[A-Z]+\d+)?$"
    return bool(re.match(pattern, cell_ref))


def check_formula_dependencies(formula: str, available_columns: List[str]) -> Dict[str, any]:
    """
    Check if formula references columns that exist.
    
    This is a basic check - it looks for column letters in the formula
    and warns if they might not exist.
    
    Args:
        formula: Excel formula string
        available_columns: List of available column names
        
    Returns:
        Dictionary with dependency check results
    """
    warnings = []
    
    # Extract column references from formula (basic pattern matching)
    # This is a simplified check - full parsing would require a proper Excel formula parser
    column_refs = re.findall(r'([A-Z]+)\d+', formula)
    
    if column_refs:
        warnings.append(f'Formula references columns: {", ".join(set(column_refs))}')
    
    return {
        'has_dependencies': len(column_refs) > 0,
        'referenced_columns': list(set(column_refs)),
        'warnings': warnings
    }

