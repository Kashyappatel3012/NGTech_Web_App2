"""
Flask Middleware for Automatic Activity Logging
This module provides decorators and middleware to automatically log user activities
"""
from functools import wraps
from flask import request, session, current_app
from flask_login import current_user
from Admin_Dashboard_Files.user_activity_logger import log_user_activity
import json

def log_activity(activity_type=None, activity_description=None):
    """
    Decorator to log user activity for a specific route
    
    Usage:
        @app.route('/some_route')
        @login_required
        @log_activity(activity_type='data_access', activity_description='Accessed some data')
        def some_route():
            ...
    """
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            # Execute the function first
            response = f(*args, **kwargs)
            
            # Log the activity
            try:
                # Get user info
                user_id = current_user.id if current_user.is_authenticated else None
                username = current_user.username if current_user.is_authenticated else None
                employee_name = current_user.employee_name if current_user.is_authenticated else None
                department = current_user.department if current_user.is_authenticated else None
                
                # Get request info
                ip_address = request.remote_addr
                request_method = request.method
                request_url = request.url
                user_agent = request.headers.get('User-Agent', '')
                session_id = session.get('session_id', session.get('_id', ''))
                
                # Get request data (for POST/PUT requests)
                request_data = None
                if request_method in ['POST', 'PUT', 'PATCH']:
                    try:
                        if request.is_json:
                            request_data = request.get_json()
                        elif request.form:
                            request_data = dict(request.form)
                        elif request.data:
                            try:
                                request_data = json.loads(request.data.decode('utf-8'))
                            except:
                                request_data = request.data.decode('utf-8', errors='ignore')
                    except:
                        pass
                
                # Get response status
                response_status = None
                if hasattr(response, 'status_code'):
                    response_status = response.status_code
                elif isinstance(response, tuple):
                    response_status = response[1] if len(response) > 1 else None
                
                # Determine activity type and description
                act_type = activity_type or f'{request_method.lower()}_request'
                act_description = activity_description or f'{request_method} request to {request.path}'
                
                # Additional details
                additional_details = {
                    'route': request.endpoint,
                    'view_args': dict(request.view_args) if request.view_args else {},
                    'query_params': dict(request.args)
                }
                
                # Log the activity
                log_user_activity(
                    user_id=user_id,
                    username=username,
                    employee_name=employee_name,
                    department=department,
                    ip_address=ip_address,
                    activity_type=act_type,
                    activity_description=act_description,
                    request_method=request_method,
                    request_url=request_url,
                    request_data=request_data,
                    response_status=response_status,
                    session_id=session_id,
                    user_agent=user_agent,
                    additional_details=additional_details
                )
                
            except Exception as e:
                # Don't fail the request if logging fails
                print(f"Error logging activity: {e}")
                import traceback
                traceback.print_exc()
            
            return response
        return decorated_function
    return decorator

def log_activity_after_request():
    """
    Function to be called after each request to log activity
    This can be registered with Flask's after_request hook
    """
    def log_request(response):
        try:
            # Only log if user is authenticated
            if not current_user.is_authenticated:
                return response
            
            # Skip logging for static files and certain routes
            skip_paths = ['/static/', '/favicon.ico', '/admin/get_logs', '/admin/download_logs']
            if any(request.path.startswith(path) for path in skip_paths):
                return response
            
            # Get user info
            user_id = current_user.id
            username = current_user.username
            employee_name = current_user.employee_name
            department = current_user.department
            
            # Get request info
            ip_address = request.remote_addr
            request_method = request.method
            request_url = request.url
            user_agent = request.headers.get('User-Agent', '')
            session_id = session.get('session_id', session.get('_id', ''))
            
            # Get request data
            request_data = None
            if request_method in ['POST', 'PUT', 'PATCH']:
                try:
                    if request.is_json:
                        request_data = request.get_json()
                    elif request.form:
                        request_data = dict(request.form)
                except:
                    pass
            
            # Get response status
            response_status = response.status_code if hasattr(response, 'status_code') else None
            
            # Determine activity type
            activity_type = f'{request_method.lower()}_request'
            activity_description = f'{request_method} request to {request.path}'
            
            # Additional details
            additional_details = {
                'route': request.endpoint,
                'view_args': dict(request.view_args) if request.view_args else {},
                'query_params': dict(request.args)
            }
            
            # Log the activity
            log_user_activity(
                user_id=user_id,
                username=username,
                employee_name=employee_name,
                department=department,
                ip_address=ip_address,
                activity_type=activity_type,
                activity_description=activity_description,
                request_method=request_method,
                request_url=request_url,
                request_data=request_data,
                response_status=response_status,
                session_id=session_id,
                user_agent=user_agent,
                additional_details=additional_details
            )
            
        except Exception as e:
            # Don't fail the request if logging fails
            print(f"Error logging activity in after_request: {e}")
        
        return response
    
    return log_request

