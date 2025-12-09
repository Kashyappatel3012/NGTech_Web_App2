"""
Error Handler Utilities
Provides secure error handling to prevent information disclosure
"""
import logging
import traceback

# Configure logging
logger = logging.getLogger(__name__)

def get_safe_error_message(exception, include_details=False):
    """
    Get a safe error message that doesn't leak sensitive information.
    
    Args:
        exception: The exception object
        include_details: If True, include generic error type (for debugging)
    
    Returns:
        str: Safe error message
    """
    # Log the full error for server-side debugging
    logger.error(f"Error occurred: {type(exception).__name__}: {str(exception)}", exc_info=True)
    
    # Return generic error message to client
    if include_details:
        # Only include exception type, not the message
        return f"An error occurred: {type(exception).__name__}"
    else:
        return "An error occurred. Please try again later."

def handle_exception_safely(exception, log_context=""):
    """
    Handle exception safely, logging full details but returning safe message.
    
    Args:
        exception: The exception object
        log_context: Additional context for logging
    
    Returns:
        str: Safe error message for client
    """
    # Log full error with context
    error_msg = f"{log_context} - {type(exception).__name__}: {str(exception)}"
    logger.error(error_msg, exc_info=True)
    
    # Return generic message
    return "An error occurred. Please try again later."

def sanitize_error_response(error_message):
    """
    Sanitize error message before sending to client.
    Removes sensitive information like file paths, stack traces, etc.
    
    Args:
        error_message: Original error message
    
    Returns:
        str: Sanitized error message
    """
    if not error_message:
        return "An error occurred. Please try again later."
    
    # Remove common sensitive patterns
    import re
    
    # Remove file paths
    error_message = re.sub(r'[A-Z]:\\[^\s]+|/[^\s]+', '[path removed]', error_message)
    
    # Remove stack trace indicators
    error_message = re.sub(r'Traceback.*?File.*?line.*?\n', '', error_message, flags=re.DOTALL)
    
    # Remove exception type details that might leak info
    sensitive_patterns = [
        r'password',
        r'secret',
        r'key',
        r'token',
        r'credential',
        r'api_key',
        r'access_token',
        r'database',
        r'connection',
        r'query',
        r'sql',
    ]
    
    for pattern in sensitive_patterns:
        if re.search(pattern, error_message, re.IGNORECASE):
            return "An error occurred. Please try again later."
    
    return error_message

