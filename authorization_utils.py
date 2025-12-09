"""
Authorization Utilities
Provides functions for access control and privilege checks
"""
from flask import abort
from flask_login import current_user

def require_department(*allowed_departments):
    """
    Decorator to require user to be in one of the allowed departments.
    
    Args:
        *allowed_departments: Variable number of department names (e.g., "HR", "Admin", "Audit")
    
    Returns:
        function: Decorator function
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                abort(401)  # Unauthorized
            
            if current_user.department not in allowed_departments:
                abort(403)  # Forbidden
            
            return func(*args, **kwargs)
        wrapper.__name__ = func.__name__
        return wrapper
    return decorator

def check_user_access(user_id, allow_self=True, allowed_departments=None):
    """
    Check if current user can access data for a specific user_id.
    Implements horizontal privilege control.
    
    Args:
        user_id: ID of the user whose data is being accessed
        allow_self: If True, allow users to access their own data
        allowed_departments: List of departments that can access any user's data (e.g., ["HR", "Admin"])
    
    Returns:
        bool: True if access is allowed, False otherwise
    
    Raises:
        HTTPException: 403 if access is denied
    """
    if not current_user.is_authenticated:
        abort(401)  # Unauthorized
    
    # If user is accessing their own data and self-access is allowed
    if allow_self and current_user.id == user_id:
        return True
    
    # If user's department is in allowed departments (vertical privilege)
    if allowed_departments and current_user.department in allowed_departments:
        return True
    
    # Access denied
    abort(403)  # Forbidden

def require_hr():
    """Require user to be in HR department"""
    if not current_user.is_authenticated:
        abort(401)
    if current_user.department != "HR":
        abort(403)

def require_admin():
    """Require user to be in Admin department"""
    if not current_user.is_authenticated:
        abort(401)
    if current_user.department != "Admin":
        abort(403)

def require_audit():
    """Require user to be in Audit department"""
    if not current_user.is_authenticated:
        abort(401)
    if current_user.department != "Audit":
        abort(403)

def require_department_list(*departments):
    """Require user to be in one of the specified departments"""
    if not current_user.is_authenticated:
        abort(401)
    if current_user.department not in departments:
        abort(403)

def can_access_user_data(target_user_id):
    """
    Check if current user can access another user's data.
    Returns True if:
    - User is accessing their own data, OR
    - User is HR or Admin (can access any user's data)
    
    Args:
        target_user_id: ID of the user whose data is being accessed
    
    Returns:
        bool: True if access is allowed
    """
    if not current_user.is_authenticated:
        return False
    
    # Users can always access their own data
    if current_user.id == target_user_id:
        return True
    
    # HR and Admin can access any user's data
    if current_user.department in ["HR", "Admin"]:
        return True
    
    return False

def can_modify_user_data(target_user_id):
    """
    Check if current user can modify another user's data.
    Returns True if:
    - User is HR (can modify any user's data)
    
    Args:
        target_user_id: ID of the user whose data is being modified
    
    Returns:
        bool: True if modification is allowed
    """
    if not current_user.is_authenticated:
        return False
    
    # Only HR can modify user data
    if current_user.department == "HR":
        return True
    
    return False

