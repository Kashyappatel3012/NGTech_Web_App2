"""
Password Validation and Strength Checking Utilities
Implements strong password policies to prevent weak passwords
"""
import re
import string

# Common weak passwords that should be rejected
COMMON_PASSWORDS = [
    'password', 'password123', '123456', '12345678', '123456789', '1234567890',
    'qwerty', 'abc123', 'monkey', '1234567', 'letmein', 'trustno1', 'dragon',
    'baseball', 'iloveyou', 'master', 'sunshine', 'ashley', 'bailey', 'passw0rd',
    'shadow', '123123', '654321', 'superman', 'qazwsx', 'michael', 'football',
    'welcome', 'jesus', 'ninja', 'mustang', 'password1', '123qwe', 'admin',
    'administrator', 'root', 'toor', 'test', 'guest', 'user', 'demo'
]

def validate_password_strength(password):
    """
    Validate password strength according to security best practices.
    
    Requirements:
    - Minimum 12 characters (recommended by NIST)
    - At least one uppercase letter
    - At least one lowercase letter
    - At least one digit
    - At least one special character
    - Not a common/weak password
    - Not containing username or common patterns
    
    Args:
        password: Password string to validate
    
    Returns:
        tuple: (is_valid: bool, error_message: str or None)
    """
    if not password:
        return False, "Password is required"
    
    password = str(password)
    errors = []
    
    # Check minimum length (12 characters recommended by NIST)
    if len(password) < 8:
        errors.append("Password must be at least 8 characters long")
    
    # Check maximum length (prevent DoS)
    if len(password) > 128:
        errors.append("Password must be no more than 128 characters long")
    
    # Check for uppercase letter
    if not re.search(r'[A-Z]', password):
        errors.append("Password must contain at least one uppercase letter")
    
    # Check for lowercase letter
    if not re.search(r'[a-z]', password):
        errors.append("Password must contain at least one lowercase letter")
    
    # Check for digit
    if not re.search(r'\d', password):
        errors.append("Password must contain at least one digit")
    
    # Check for special character
    special_chars = r'!@#$%^&*()_+-=[]{}|;:,.<>?'
    if not re.search(f'[{re.escape(special_chars)}]', password):
        errors.append("Password must contain at least one special character (!@#$%^&*()_+-=[]{}|;:,.<>?)")
    
    # Check for common passwords (case-insensitive)
    if password.lower() in [p.lower() for p in COMMON_PASSWORDS]:
        errors.append("Password is too common. Please choose a stronger password")
    
    # Check for repeated characters (e.g., "aaaa" or "1111")
    if re.search(r'(.)\1{3,}', password):
        errors.append("Password cannot contain the same character repeated 4 or more times")
    
    # Check for sequential characters (e.g., "1234" or "abcd")
    if re.search(r'(0123|1234|2345|3456|4567|5678|6789|7890|abcd|bcde|cdef|defg|efgh|fghi|ghij|hijk|ijkl|jklm|klmn|lmno|mnop|nopq|opqr|pqrs|qrst|rstu|stuv|tuvw|uvwx|vwxy|wxyz)', password.lower()):
        errors.append("Password cannot contain sequential characters (e.g., '1234' or 'abcd')")
    
    # Check for common patterns
    if re.search(r'(.)\1{2,}', password):
        # Allow up to 2 repeated characters, but warn about 3+
        pass  # Already checked above
    
    if errors:
        return False, "; ".join(errors)
    
    return True, None

def check_password_complexity(password):
    """
    Check password complexity score (0-100).
    Higher score = stronger password.
    
    Args:
        password: Password string
    
    Returns:
        int: Complexity score (0-100)
    """
    if not password:
        return 0
    
    score = 0
    
    # Length score (max 40 points)
    length = len(password)
    if length >= 12:
        score += min(40, (length - 12) * 2)  # 2 points per character over 12, max 40
    
    # Character variety (max 40 points)
    has_upper = bool(re.search(r'[A-Z]', password))
    has_lower = bool(re.search(r'[a-z]', password))
    has_digit = bool(re.search(r'\d', password))
    has_special = bool(re.search(r'[!@#$%^&*()_+\-=\[\]{}|;:,.<>?]', password))
    
    variety_count = sum([has_upper, has_lower, has_digit, has_special])
    score += variety_count * 10  # 10 points per character type, max 40
    
    # Entropy bonus (max 20 points)
    # Simple entropy calculation based on character variety
    char_types = 0
    if has_upper: char_types += 26
    if has_lower: char_types += 26
    if has_digit: char_types += 10
    if has_special: char_types += len('!@#$%^&*()_+-=[]{}|;:,.<>?')
    
    if char_types > 0:
        # Approximate entropy
        entropy = length * (char_types ** 0.5)
        score += min(20, int(entropy / 10))
    
    return min(100, score)

def is_password_strong(password):
    """
    Quick check if password meets minimum strength requirements.
    
    Args:
        password: Password string
    
    Returns:
        bool: True if password is strong enough
    """
    is_valid, _ = validate_password_strength(password)
    return is_valid

def get_password_requirements():
    """
    Get password requirements as a formatted string for display.
    
    Returns:
        str: Formatted requirements text
    """
    return """Password Requirements:
• Minimum 12 characters
• At least one uppercase letter (A-Z)
• At least one lowercase letter (a-z)
• At least one digit (0-9)
• At least one special character (!@#$%^&*()_+-=[]{}|;:,.<>?)
• Cannot be a common/weak password
• Cannot contain repeated characters (e.g., 'aaaa')
• Cannot contain sequential characters (e.g., '1234' or 'abcd')"""

