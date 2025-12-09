"""
Rate Limiting Utilities
Prevents brute force attacks and credential stuffing
"""
from datetime import datetime, timedelta
from collections import defaultdict
import time

# In-memory storage for rate limiting (use Redis in production)
# Format: {key: [(timestamp1, count1), (timestamp2, count2), ...]}
_rate_limit_store = defaultdict(list)

# Rate limit configurations
RATE_LIMITS = {
    'login_per_ip': {
        'max_attempts': 5,  # Max 5 login attempts
        'window': timedelta(minutes=15)  # Per 15 minutes
    },
    'login_per_username': {
        'max_attempts': 5,  # Max 5 login attempts
        'window': timedelta(minutes=15)  # Per 15 minutes
    },
    'otp_verification_per_ip': {
        'max_attempts': 10,  # Max 10 OTP attempts
        'window': timedelta(minutes=15)  # Per 15 minutes
    },
    'credential_stuffing_detection': {
        'max_usernames': 10,  # Max 10 different usernames
        'window': timedelta(minutes=15)  # Per 15 minutes from same IP
    }
}

def check_rate_limit(limit_type, identifier):
    """
    Check if rate limit has been exceeded.
    
    Args:
        limit_type: Type of rate limit (e.g., 'login_per_ip', 'login_per_username')
        identifier: Unique identifier (IP address, username, etc.)
    
    Returns:
        tuple: (is_allowed: bool, remaining_attempts: int, reset_time: datetime or None)
    """
    if limit_type not in RATE_LIMITS:
        return True, -1, None
    
    config = RATE_LIMITS[limit_type]
    key = f"{limit_type}:{identifier}"
    current_time = datetime.now()
    window_start = current_time - config['window']
    
    # Clean old entries
    rate_list = _rate_limit_store[key]
    _rate_limit_store[key] = [
        (ts, count) for ts, count in rate_list 
        if ts > window_start
    ]
    
    # Count attempts in current window
    total_attempts = sum(count for _, count in _rate_limit_store[key])
    
    # Check if limit exceeded
    if total_attempts >= config['max_attempts']:
        # Find oldest entry to calculate reset time
        if _rate_limit_store[key]:
            oldest_time = min(ts for ts, _ in _rate_limit_store[key])
            reset_time = oldest_time + config['window']
        else:
            reset_time = current_time + config['window']
        
        remaining = 0
        return False, remaining, reset_time
    
    remaining = config['max_attempts'] - total_attempts
    return True, remaining, None

def record_attempt(limit_type, identifier):
    """
    Record an attempt for rate limiting.
    
    Args:
        limit_type: Type of rate limit
        identifier: Unique identifier (IP address, username, etc.)
    """
    if limit_type not in RATE_LIMITS:
        return
    
    key = f"{limit_type}:{identifier}"
    current_time = datetime.now()
    
    # Add attempt
    _rate_limit_store[key].append((current_time, 1))
    
    # Clean old entries periodically (keep last 100 entries per key)
    if len(_rate_limit_store[key]) > 100:
        window_start = current_time - RATE_LIMITS[limit_type]['window']
        _rate_limit_store[key] = [
            (ts, count) for ts, count in _rate_limit_store[key] 
            if ts > window_start
        ]

def check_credential_stuffing(ip_address, username):
    """
    Detect credential stuffing attacks.
    Checks if same IP is trying multiple different usernames.
    
    Args:
        ip_address: IP address of the request
        username: Username being attempted
    
    Returns:
        tuple: (is_suspicious: bool, message: str or None)
    """
    key = f"credential_stuffing:{ip_address}"
    current_time = datetime.now()
    window_start = current_time - RATE_LIMITS['credential_stuffing_detection']['window']
    
    # Get unique usernames attempted from this IP
    if key not in _rate_limit_store:
        _rate_limit_store[key] = []
    
    # Clean old entries
    _rate_limit_store[key] = [
        (ts, uname) for ts, uname in _rate_limit_store[key] 
        if ts > window_start
    ]
    
    # Get unique usernames
    unique_usernames = set(uname for _, uname in _rate_limit_store[key])
    
    # Add current username
    unique_usernames.add(username)
    _rate_limit_store[key].append((current_time, username))
    
    # Check if too many different usernames
    if len(unique_usernames) > RATE_LIMITS['credential_stuffing_detection']['max_usernames']:
        return True, f"Too many different usernames attempted from this IP. Please try again later."
    
    return False, None

def reset_rate_limit(limit_type, identifier):
    """
    Reset rate limit for a specific identifier.
    Useful after successful authentication.
    
    Args:
        limit_type: Type of rate limit
        identifier: Unique identifier
    """
    key = f"{limit_type}:{identifier}"
    if key in _rate_limit_store:
        del _rate_limit_store[key]

def get_rate_limit_status(limit_type, identifier):
    """
    Get current rate limit status without recording an attempt.
    
    Args:
        limit_type: Type of rate limit
        identifier: Unique identifier
    
    Returns:
        dict: Status information
    """
    is_allowed, remaining, reset_time = check_rate_limit(limit_type, identifier)
    
    return {
        'allowed': is_allowed,
        'remaining_attempts': remaining,
        'reset_time': reset_time.isoformat() if reset_time else None,
        'max_attempts': RATE_LIMITS[limit_type]['max_attempts'] if limit_type in RATE_LIMITS else 0
    }

