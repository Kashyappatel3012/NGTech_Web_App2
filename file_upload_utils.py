"""
File Upload Security Utilities
Prevents unrestricted file upload and validates file content
"""
import os
from werkzeug.utils import secure_filename
from security_utils import validate_file_extension, sanitize_filename

# MIME type mappings for allowed file types
ALLOWED_MIME_TYPES = {
    'image/png': ['.png'],
    'image/jpeg': ['.jpg', '.jpeg'],
    'image/gif': ['.gif'],
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
    'application/vnd.ms-excel': ['.xls'],
    'application/zip': ['.zip'],
}

def validate_file_content(file, allowed_extensions):
    """
    Validate file content using MIME type detection.
    Prevents file type spoofing by checking actual file content.
    
    Args:
        file: File object from request.files
        allowed_extensions: Set of allowed extensions (e.g., {'png', 'jpg'})
    
    Returns:
        tuple: (is_valid: bool, error_message: str)
    """
    if not file or not file.filename:
        return False, "No file provided"
    
    # Check file extension first
    if not validate_file_extension(file.filename, allowed_extensions):
        return False, f"File extension not allowed. Allowed: {', '.join(allowed_extensions)}"
    
    # Read first few bytes for MIME type detection
    try:
        file.seek(0)
        file_header = file.read(1024)  # Read first 1KB
        file.seek(0)  # Reset file pointer
        
        # Get file extension
        ext = os.path.splitext(file.filename)[1].lower()
        
        # Validate MIME type matches extension
        # For images, check magic bytes
        if ext in ['.png', '.jpg', '.jpeg', '.gif']:
            if ext == '.png' and not file_header.startswith(b'\x89PNG\r\n\x1a\n'):
                return False, "File content does not match PNG format"
            elif ext in ['.jpg', '.jpeg'] and not file_header.startswith(b'\xff\xd8\xff'):
                return False, "File content does not match JPEG format"
            elif ext == '.gif' and not file_header.startswith(b'GIF89a') and not file_header.startswith(b'GIF87a'):
                return False, "File content does not match GIF format"
        
        # For Excel files, check magic bytes
        elif ext in ['.xlsx', '.xls']:
            # XLSX is a ZIP file
            if ext == '.xlsx' and not file_header.startswith(b'PK\x03\x04'):
                return False, "File content does not match XLSX format"
            # XLS has OLE header
            elif ext == '.xls' and not file_header.startswith(b'\xd0\xcf\x11\xe0'):
                return False, "File content does not match XLS format"
        
        # For ZIP files
        elif ext == '.zip':
            if not file_header.startswith(b'PK\x03\x04'):
                return False, "File content does not match ZIP format"
        
        return True, "File content validated"
    
    except Exception as e:
        return False, f"Error validating file content: {str(e)}"

def validate_file_size(file, max_size_mb=16):
    """
    Validate file size.
    
    Args:
        file: File object from request.files
        max_size_mb: Maximum file size in MB
    
    Returns:
        tuple: (is_valid: bool, error_message: str, file_size: int)
    """
    if not file:
        return False, "No file provided", 0
    
    try:
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)  # Reset file pointer
        
        max_size_bytes = max_size_mb * 1024 * 1024
        
        if file_size > max_size_bytes:
            return False, f"File size exceeds maximum allowed size of {max_size_mb}MB", file_size
        
        return True, "File size validated", file_size
    
    except Exception as e:
        return False, f"Error checking file size: {str(e)}", 0

def secure_file_upload(file, upload_folder, allowed_extensions, max_size_mb=16, custom_filename=None):
    """
    Securely handle file upload with comprehensive validation.
    
    Args:
        file: File object from request.files
        upload_folder: Directory to save file
        allowed_extensions: Set of allowed extensions
        max_size_mb: Maximum file size in MB
        custom_filename: Optional custom filename (will be sanitized)
    
    Returns:
        tuple: (success: bool, filename: str or None, error_message: str)
    """
    # Validate file exists
    if not file or not file.filename:
        return False, None, "No file provided"
    
    # Validate file extension
    if not validate_file_extension(file.filename, allowed_extensions):
        return False, None, f"File extension not allowed. Allowed: {', '.join(allowed_extensions)}"
    
    # Validate file size
    size_valid, size_msg, file_size = validate_file_size(file, max_size_mb)
    if not size_valid:
        return False, None, size_msg
    
    # Validate file content
    content_valid, content_msg = validate_file_content(file, allowed_extensions)
    if not content_valid:
        return False, None, content_msg
    
    # Sanitize filename
    if custom_filename:
        safe_filename = sanitize_filename(custom_filename)
    else:
        safe_filename = sanitize_filename(file.filename)
    
    # Ensure upload folder exists
    os.makedirs(upload_folder, mode=0o755, exist_ok=True)
    
    # Construct full path
    file_path = os.path.join(upload_folder, safe_filename)
    
    # Ensure path is within upload folder (prevent path traversal)
    upload_folder_abs = os.path.abspath(upload_folder)
    file_path_abs = os.path.abspath(file_path)
    
    if not file_path_abs.startswith(upload_folder_abs):
        return False, None, "Invalid file path detected"
    
    try:
        # Save file
        file.save(file_path)
        return True, safe_filename, "File uploaded successfully"
    
    except Exception as e:
        return False, None, f"Error saving file: {str(e)}"

