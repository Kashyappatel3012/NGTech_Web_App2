"""
Security Utilities for Input Sanitization and Validation
Prevents various injection attacks
"""
import re
import os
from werkzeug.utils import secure_filename
from urllib.parse import urlparse
from markupsafe import Markup, escape

def sanitize_email_header(value):
    """
    Sanitize email header values to prevent email header injection.
    Removes newlines, carriage returns, and other dangerous characters.
    
    Args:
        value: String to sanitize
    
    Returns:
        str: Sanitized string safe for email headers
    """
    if not value:
        return ""
    
    # Remove newlines, carriage returns, and null bytes
    sanitized = re.sub(r'[\r\n\x00]', '', str(value))
    
    # Remove other potentially dangerous characters for email headers
    # Remove characters that could be used for header injection
    sanitized = re.sub(r'[<>]', '', sanitized)
    
    # Limit length to prevent buffer overflow
    if len(sanitized) > 2000:
        sanitized = sanitized[:2000]
    
    return sanitized.strip()

def sanitize_email_content(value):
    """
    Sanitize email content/body to prevent injection.
    Allows HTML but escapes dangerous content.
    
    Args:
        value: String to sanitize
    
    Returns:
        str: Sanitized string safe for email content
    """
    if not value:
        return ""
    
    # Remove null bytes
    sanitized = re.sub(r'\x00', '', str(value))
    
    # Limit length to prevent DoS
    if len(sanitized) > 100000:  # 100KB limit
        sanitized = sanitized[:100000]
    
    return sanitized

def sanitize_filename(filename):
    """
    Sanitize filename to prevent path traversal and other file-based attacks.
    Uses secure_filename and additional checks.
    
    Args:
        filename: Original filename
    
    Returns:
        str: Sanitized filename safe for file operations
    """
    if not filename:
        return "file"
    
    # Use werkzeug's secure_filename
    safe_name = secure_filename(filename)
    
    # Additional check: ensure no path traversal attempts
    if '..' in safe_name or '/' in safe_name or '\\' in safe_name:
        # Extract just the basename
        safe_name = os.path.basename(safe_name)
        safe_name = secure_filename(safe_name)
    
    # If still empty after sanitization, use default
    if not safe_name:
        safe_name = "file"
    
    # Limit length
    if len(safe_name) > 255:
        name, ext = os.path.splitext(safe_name)
        safe_name = name[:250] + ext
    
    return safe_name

def validate_url(url, allowed_hosts=None):
    """
    Validate URL to prevent open redirect and SSRF attacks.
    
    Args:
        url: URL to validate
        allowed_hosts: List of allowed hostnames (optional)
    
    Returns:
        str or None: Validated URL or None if invalid
    """
    if not url:
        return None
    
    try:
        parsed = urlparse(url)
        
        # Reject dangerous schemes
        if parsed.scheme not in ['http', 'https', '']:
            return None
        
        # If allowed_hosts specified, check host
        if allowed_hosts:
            if parsed.netloc not in allowed_hosts:
                return None
        
        # Reject localhost/internal IPs (SSRF protection)
        host = parsed.hostname or ''
        if host in ['localhost', '127.0.0.1', '0.0.0.0']:
            return None
        
        # Check for private IP ranges
        if host.startswith('192.168.') or host.startswith('10.') or host.startswith('172.'):
            return None
        
        return url
    except:
        return None

def sanitize_referrer(referrer):
    """
    Sanitize HTTP referrer to prevent open redirect attacks.
    
    Args:
        referrer: Referrer URL from request
    
    Returns:
        str or None: Safe referrer URL or None
    """
    if not referrer:
        return None
    
    # Only allow relative URLs or same-origin URLs
    try:
        parsed = urlparse(referrer)
        
        # If no scheme, it's relative - safe
        if not parsed.scheme:
            # Ensure it's a relative path (starts with /)
            if referrer.startswith('/'):
                return referrer
            return None
        
        # If has scheme, must be http/https and same origin
        if parsed.scheme not in ['http', 'https']:
            return None
        
        # For now, reject external referrers to prevent open redirect
        # In production, you might want to whitelist specific domains
        return None
    except:
        return None

def sanitize_for_xss(value):
    """
    Sanitize value for XSS prevention.
    Jinja2 auto-escapes by default, but this provides additional safety.
    
    Args:
        value: String to sanitize
    
    Returns:
        Markup: Safe markup object
    """
    if value is None:
        return ""
    
    # Escape HTML special characters
    return escape(str(value))

def sanitize_sql_input(value):
    """
    Validate input for SQL injection prevention.
    Note: This is a secondary check - always use parameterized queries!
    
    Args:
        value: Input value
    
    Returns:
        str: Sanitized value
    """
    if value is None:
        return None
    
    value_str = str(value)
    
    # Remove SQL comment characters
    value_str = value_str.replace('--', '')
    value_str = value_str.replace('/*', '')
    value_str = value_str.replace('*/', '')
    
    # Remove semicolons (statement terminators)
    value_str = value_str.replace(';', '')
    
    return value_str

def validate_file_extension(filename, allowed_extensions):
    """
    Validate file extension against allowed list.
    
    Args:
        filename: Filename to check
        allowed_extensions: Set of allowed extensions (e.g., {'jpg', 'png'})
    
    Returns:
        bool: True if extension is allowed
    """
    if not filename:
        return False
    
    # Get extension
    ext = os.path.splitext(filename)[1].lower().lstrip('.')
    
    return ext in allowed_extensions

def sanitize_path(path):
    """
    Sanitize file path to prevent directory traversal.
    
    Args:
        path: File path to sanitize
    
    Returns:
        str: Sanitized path
    """
    if not path:
        return ""
    
    # Normalize path
    normalized = os.path.normpath(path)
    
    # Remove any remaining path traversal attempts
    normalized = normalized.replace('..', '')
    normalized = normalized.replace('//', '/')
    normalized = normalized.replace('\\\\', '\\')
    
    # Get basename to prevent directory traversal
    return os.path.basename(normalized)

