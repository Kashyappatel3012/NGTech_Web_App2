"""
CSRF Protection Utilities
Implements CSRF token generation and validation
"""
import secrets
from functools import wraps
from flask import session, request, abort

def generate_csrf_token():
    """Generate a CSRF token and store it in session"""
    if 'csrf_token' not in session:
        session['csrf_token'] = secrets.token_urlsafe(32)
    return session['csrf_token']

def validate_csrf_token(token=None):
    """Validate CSRF token from request"""
    if token is None:
        # Try to get token from form or header
        token = request.form.get('csrf_token') or request.headers.get('X-CSRF-Token')
    
    session_token = session.get('csrf_token')
    if not token or not session_token:
        return False
    
    return secrets.compare_digest(token, session_token)

def csrf_protect(f):
    """Decorator to protect routes from CSRF attacks"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if request.method in ['POST', 'PUT', 'DELETE', 'PATCH']:
            if not validate_csrf_token():
                abort(403)  # Forbidden - CSRF token missing or invalid
        return f(*args, **kwargs)
    return decorated_function

