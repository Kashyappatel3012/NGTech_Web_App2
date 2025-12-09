"""
Excel Security Utilities
Prevents CSV/Excel injection attacks by sanitizing cell values
"""
import re

def sanitize_excel_value(value):
    """
    Sanitize value before writing to Excel cell to prevent formula injection.
    
    Excel/CSV injection occurs when cell values start with:
    - = (formula)
    - + (formula)
    - - (formula)
    - @ (formula in older Excel)
    - \t (tab character, can be used for CSV injection)
    - \r (carriage return)
    
    Args:
        value: Value to sanitize (can be string, int, float, None)
    
    Returns:
        str: Sanitized value safe for Excel cells
    """
    if value is None:
        return ""
    
    # Convert to string
    value_str = str(value)
    
    # Remove null bytes
    value_str = value_str.replace('\x00', '')
    
    # Check for formula injection patterns
    dangerous_prefixes = ['=', '+', '-', '@', '\t', '\r']
    
    # If value starts with dangerous prefix, prepend with single quote
    # Single quote in Excel forces text interpretation
    if value_str and value_str[0] in dangerous_prefixes:
        # Prepend with single quote to force text interpretation
        value_str = "'" + value_str
    
    # Also check for common formula functions that could be dangerous
    dangerous_functions = [
        'cmd', 'powershell', 'rundll32', 'regsvr32',
        'HYPERLINK', 'IMPORTXML', 'WEBSERVICE', 'FILTERXML'
    ]
    
    # Check if value contains dangerous function calls
    value_upper = value_str.upper()
    for func in dangerous_functions:
        if func in value_upper:
            # Prepend with single quote to force text interpretation
            if not value_str.startswith("'"):
                value_str = "'" + value_str
            break
    
    # Limit length to prevent DoS
    if len(value_str) > 32767:  # Excel cell character limit
        value_str = value_str[:32767]
    
    return value_str

def sanitize_excel_cell_value(value):
    """
    Alias for sanitize_excel_value for backward compatibility.
    """
    return sanitize_excel_value(value)

def is_safe_excel_value(value):
    """
    Check if a value is safe to write to Excel without sanitization.
    
    Args:
        value: Value to check
    
    Returns:
        bool: True if safe, False if needs sanitization
    """
    if value is None:
        return True
    
    value_str = str(value)
    
    # Check for dangerous prefixes
    if value_str and value_str[0] in ['=', '+', '-', '@', '\t', '\r']:
        return False
    
    # Check for dangerous functions
    dangerous_functions = [
        'cmd', 'powershell', 'rundll32', 'regsvr32',
        'HYPERLINK', 'IMPORTXML', 'WEBSERVICE', 'FILTERXML'
    ]
    
    value_upper = value_str.upper()
    for func in dangerous_functions:
        if func in value_upper:
            return False
    
    return True

