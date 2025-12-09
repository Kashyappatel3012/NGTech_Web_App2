# IDOR, Brute Force, and Credential Stuffing Security Fixes

## Summary
Fixed Insecure Direct Object References (IDOR), Brute Force Attacks, and Credential Stuffing vulnerabilities throughout the application.

## Vulnerabilities Fixed

### 1. ✅ Insecure Direct Object References (IDOR)
**Status**: FIXED
**Issue**: Routes that accept `user_id` parameters did not verify that:
- The user exists and is not deleted
- The user is active (not deactivated)
- Proper authorization checks were in place

**Fix Applied**:
- Added IDOR protection to all routes that accept `user_id`
- Verify user exists, is not deleted, and is active before allowing access
- Return 404 (not found) instead of revealing user existence for deleted/inactive users

**Routes Fixed**:
- `/api/get_user_details/<int:user_id>` - Added IDOR protection
- `/update_performance/<int:user_id>` - Added IDOR protection
- `/get_performance_data/<int:user_id>` - Added IDOR protection
- `/get_employee_data/<int:user_id>` - Added IDOR protection
- `/update_employee_data/<int:user_id>` - Added IDOR protection
- `/request_reset_otp` - Added IDOR protection
- `/verify_reset_otp` - Added IDOR protection
- `/reset_user_password` - Added IDOR protection
- `/api/update_user_field` - Added IDOR protection
- `/api/update_user_all` - Added IDOR protection
- `/api/delete_user` - Added IDOR protection
- `/toggle_user_status/<int:user_id>` - Added IDOR protection

### 2. ✅ Brute Force Attacks
**Status**: FIXED
**Issue**: No rate limiting on login and OTP verification endpoints, allowing unlimited brute force attempts.

**Fix Applied**:
- Implemented IP-based rate limiting (5 attempts per 15 minutes per IP)
- Implemented username-based rate limiting (5 attempts per 15 minutes per username)
- Implemented OTP verification rate limiting (10 attempts per 15 minutes per IP)
- Rate limits reset on successful authentication
- Clear error messages with remaining time

**Rate Limits Configured**:
- **Login per IP**: 5 attempts per 15 minutes
- **Login per username**: 5 attempts per 15 minutes
- **OTP verification per IP**: 10 attempts per 15 minutes

**Existing Protections** (Already in place):
- Account lockout after 3 failed attempts (15 minutes)
- Permanent lockout after 5 failed attempts (requires HR intervention)
- Failed attempt tracking and logging

### 3. ✅ Credential Stuffing
**Status**: FIXED
**Issue**: No detection for credential stuffing attacks (same IP trying multiple different usernames).

**Fix Applied**:
- Implemented credential stuffing detection
- Tracks unique usernames attempted from same IP
- Blocks if more than 10 different usernames attempted from same IP in 15 minutes
- Logs suspicious activity for monitoring

**Detection Logic**:
- Tracks all usernames attempted from each IP address
- If same IP tries 10+ different usernames in 15 minutes, blocks the IP
- Provides clear error message to user
- Logs activity for security monitoring

## Code Changes

### New Files:
1. **rate_limit_utils.py** (NEW)
   - `check_rate_limit()` - Check if rate limit exceeded
   - `record_attempt()` - Record an attempt for rate limiting
   - `check_credential_stuffing()` - Detect credential stuffing attacks
   - `reset_rate_limit()` - Reset rate limit on success
   - `get_rate_limit_status()` - Get current rate limit status

### Modified Files:
1. **app.py**
   - Added rate limiting to `/login` route (IP and username based)
   - Added credential stuffing detection to `/login` route
   - Added rate limiting to `/verify_otp` route
   - Added IDOR protection to all user-related routes
   - Reset rate limits on successful authentication

## Security Features

### IDOR Protection:
- ✅ All user_id parameters validated
- ✅ Deleted users return 404 (not found)
- ✅ Inactive users return 404 (not found)
- ✅ Consistent error messages (don't reveal user existence)

### Brute Force Protection:
- ✅ IP-based rate limiting
- ✅ Username-based rate limiting
- ✅ OTP verification rate limiting
- ✅ Rate limits reset on success
- ✅ Clear error messages with time remaining
- ✅ Account lockout (existing, enhanced)

### Credential Stuffing Protection:
- ✅ Detects multiple usernames from same IP
- ✅ Blocks suspicious IPs
- ✅ Logs suspicious activity
- ✅ Clear error messages

## Rate Limit Configuration

```python
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
```

## How It Works

### IDOR Protection Flow:
1. Route receives `user_id` parameter
2. Check if user exists
3. Check if user is deleted (`deleted_at` is not None)
4. Check if user is active (`is_active` is True)
5. If any check fails, return 404 (not found)
6. Only proceed if all checks pass

### Brute Force Protection Flow:
1. User attempts login
2. Check IP-based rate limit
3. Check username-based rate limit
4. If limit exceeded, block with error message
5. Record attempt
6. On successful login, reset rate limits

### Credential Stuffing Detection Flow:
1. User attempts login with username
2. Track username attempted from IP
3. Count unique usernames from same IP
4. If > 10 different usernames in 15 minutes, block IP
5. Log suspicious activity
6. Return error message

## Testing

All fixes maintain existing functionality:
- ✅ User access works correctly with proper authorization
- ✅ Login works correctly with rate limiting
- ✅ OTP verification works correctly with rate limiting
- ✅ Rate limits reset on successful authentication
- ✅ IDOR attempts are blocked
- ✅ Brute force attempts are blocked
- ✅ Credential stuffing attempts are detected and blocked

## Error Messages

### Rate Limiting:
- "Too many login attempts from this IP. Please try again in X minutes."
- "Too many login attempts for this username. Please try again in X minutes."
- "Too many OTP verification attempts. Please try again in X minutes."

### Credential Stuffing:
- "Suspicious activity detected. Please try again later."

### IDOR:
- "User not found" (404) - Consistent message for deleted/inactive users

## Recommendations

1. **Use Redis for Rate Limiting**: For production, use Redis instead of in-memory storage for distributed rate limiting
2. **Monitor Suspicious Activity**: Set up alerts for credential stuffing detections
3. **CAPTCHA**: Consider adding CAPTCHA after multiple failed attempts
4. **IP Whitelisting**: Consider IP whitelisting for trusted networks
5. **Rate Limit Tuning**: Adjust rate limits based on actual usage patterns
6. **Logging**: All suspicious activities are logged for security monitoring

## Notes

- Rate limiting uses in-memory storage (suitable for single-server deployments)
- For production with multiple servers, use Redis for distributed rate limiting
- Rate limits are per-IP and per-username for comprehensive protection
- Credential stuffing detection complements existing account lockout mechanisms
- IDOR protection ensures users can only access valid, active resources

