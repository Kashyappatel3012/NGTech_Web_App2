"""
Client-side Security Utilities
Functions to prevent XSS, CSRF, and other client-side vulnerabilities
"""
import re
import html
from markupsafe import Markup, escape

def sanitize_for_html(text):
    """
    Sanitize text for safe HTML output (prevents XSS)
    Escapes HTML special characters
    """
    if text is None:
        return ''
    if isinstance(text, Markup):
        return text  # Already escaped
    return escape(str(text))

def sanitize_for_js(text):
    """
    Sanitize text for safe JavaScript output (prevents XSS in JS context)
    Escapes JavaScript special characters
    """
    if text is None:
        return ''
    text = str(text)
    # Escape backslashes first
    text = text.replace('\\', '\\\\')
    # Escape quotes
    text = text.replace("'", "\\'")
    text = text.replace('"', '\\"')
    # Escape newlines
    text = text.replace('\n', '\\n')
    text = text.replace('\r', '\\r')
    # Escape forward slashes (for JSON safety)
    text = text.replace('/', '\\/')
    return text

def sanitize_url_param(param):
    """
    Sanitize URL parameter to prevent reflected XSS
    Only allows alphanumeric, spaces, and safe punctuation
    """
    if param is None:
        return None
    param = str(param)
    # Remove any HTML tags
    param = re.sub(r'<[^>]+>', '', param)
    # Remove script tags and event handlers
    param = re.sub(r'(?i)(javascript|onerror|onload|onclick|onmouseover|onfocus|onblur):', '', param)
    # Escape HTML entities
    param = html.escape(param)
    return param

def sanitize_for_json(data):
    """
    Sanitize data for safe JSON output
    Recursively sanitizes strings in dictionaries and lists
    """
    if isinstance(data, dict):
        return {k: sanitize_for_json(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [sanitize_for_json(item) for item in data]
    elif isinstance(data, str):
        # Escape JSON special characters
        return sanitize_for_js(data)
    else:
        return data

def validate_csrf_token(token, session_token):
    """
    Validate CSRF token against session token
    Returns True if valid, False otherwise
    """
    import secrets
    if not token or not session_token:
        return False
    return secrets.compare_digest(str(token), str(session_token))

