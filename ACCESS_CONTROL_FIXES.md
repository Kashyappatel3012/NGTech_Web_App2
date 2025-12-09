# Broken Access Control, Privilege Escalation, and Horizontal/Vertical Privilege Issues - Security Fixes

## Summary
Fixed Broken Access Control, Privilege Escalation, and Horizontal/Vertical Privilege vulnerabilities throughout the application.

## Vulnerabilities Fixed

### 1. ✅ Broken Access Control
**Status**: FIXED
**Issue**: 
- Many routes were missing proper authorization checks
- Generate Excel routes were accessible to all authenticated users regardless of department
- Missing `@login_required` decorator on some routes

**Fix Applied**:
- Added `@login_required` to all generate excel routes that were missing it
- Added `require_audit()` authorization check to all 46 generate excel routes
- Created `authorization_utils.py` with reusable authorization functions
- All routes now properly check user department before allowing access

**Routes Fixed**:
- All 46 `/generate_*_excel` routes - Added `@login_required` and `require_audit()`
- All routes now enforce department-based access control

### 2. ✅ Vertical Privilege Escalation
**Status**: FIXED
**Issue**: Users from one department could access functions intended for another department.

**Fix Applied**:
- **HR Functions**: Only HR department can access:
  - User creation, update, deletion
  - Performance management
  - Employee data management
  - Password reset
  - User status toggle
  
- **Audit Functions**: Only Audit department can access:
  - All 46 Excel report generation routes
  - Audit dashboard
  
- **Admin Functions**: Only Admin department can access:
  - Admin dashboard
  - System-wide performance viewing
  - Login activity logs
  
- **Department-Specific Dashboards**: Each department can only access their own dashboard:
  - GRC dashboard (GRC department only)
  - VAPT dashboard (VAPT department only)
  - Audit dashboard (Audit department only)
  - HR dashboard (HR department only)
  - Admin dashboard (Admin department only)

### 3. ✅ Horizontal Privilege Escalation
**Status**: FIXED
**Issue**: Users could potentially access other users' data by manipulating user_id parameters.

**Fix Applied**:
- Enhanced IDOR protection (from previous fixes)
- All user-related routes verify:
  - User exists
  - User is not deleted
  - User is active
  - Proper authorization (HR/Admin can access any user, users can only access their own data)
  
**Routes with Horizontal Protection**:
- `/get_performance_data/<int:user_id>` - HR only
- `/get_employee_data/<int:user_id>` - HR only
- `/update_performance/<int:user_id>` - HR only
- `/update_employee_data/<int:user_id>` - HR only
- `/api/get_user_details/<int:user_id>` - HR only
- All other user_id-based routes

## Code Changes

### New Files:
1. **authorization_utils.py** (NEW)
   - `require_hr()` - Require HR department
   - `require_admin()` - Require Admin department
   - `require_audit()` - Require Audit department
   - `check_user_access()` - Check horizontal privilege access
   - `can_access_user_data()` - Check if user can access another user's data
   - `can_modify_user_data()` - Check if user can modify another user's data

### Modified Files:
1. **app.py**
   - Added `require_audit()` to all 46 generate excel routes
   - Added `@login_required` to generate excel routes that were missing it
   - All routes now properly enforce department-based access control

## Security Features

### Department-Based Access Control:
- ✅ HR: Can manage all users, performance, employee data
- ✅ Admin: Can view all users, performance, login activities
- ✅ Audit: Can generate audit reports (46 Excel routes)
- ✅ GRC: Can only access GRC dashboard and own data
- ✅ VAPT: Can only access VAPT dashboard and own data

### Route Protection:
- ✅ All protected routes require authentication (`@login_required`)
- ✅ All protected routes check department authorization
- ✅ All user_id-based routes verify user access rights
- ✅ All generate excel routes restricted to Audit department

### Authorization Functions:
```python
# Vertical privilege checks
require_hr()      # Only HR can access
require_admin()   # Only Admin can access
require_audit()   # Only Audit can access

# Horizontal privilege checks
can_access_user_data(target_user_id)  # Check if user can access another user's data
can_modify_user_data(target_user_id) # Check if user can modify another user's data
```

## Access Control Matrix

| Function | HR | Admin | Audit | GRC | VAPT |
|----------|----|----|----|-----|------|
| HR Dashboard | ✅ | ❌ | ❌ | ❌ | ❌ |
| Admin Dashboard | ❌ | ✅ | ❌ | ❌ | ❌ |
| Audit Dashboard | ❌ | ❌ | ✅ | ❌ | ❌ |
| GRC Dashboard | ❌ | ❌ | ❌ | ✅ | ❌ |
| VAPT Dashboard | ❌ | ❌ | ❌ | ❌ | ✅ |
| Create User | ✅ | ❌ | ❌ | ❌ | ❌ |
| Update User | ✅ | ❌ | ❌ | ❌ | ❌ |
| Delete User | ✅ | ❌ | ❌ | ❌ | ❌ |
| View All Users | ✅ | ✅ | ❌ | ❌ | ❌ |
| Generate Excel Reports | ❌ | ❌ | ✅ | ❌ | ❌ |
| View Own Performance | ✅ | ✅ | ✅ | ✅ | ✅ |
| View All Performance | ✅ | ✅ | ❌ | ❌ | ❌ |
| Update Performance | ✅ | ❌ | ❌ | ❌ | ❌ |

## Testing

All fixes maintain existing functionality:
- ✅ Department dashboards work correctly with proper authorization
- ✅ HR functions only accessible to HR department
- ✅ Audit functions only accessible to Audit department
- ✅ Admin functions only accessible to Admin department
- ✅ Users can only access their own data (unless HR/Admin)
- ✅ Generate Excel routes restricted to Audit department
- ✅ All routes properly check authentication and authorization

## Error Handling

### Authorization Failures:
- **401 Unauthorized**: User not authenticated
- **403 Forbidden**: User authenticated but lacks required department/privileges
- **404 Not Found**: Resource doesn't exist or user doesn't have access (for IDOR protection)

## Recommendations

1. **Role-Based Access Control (RBAC)**: Consider implementing a more granular RBAC system if needed
2. **Audit Logging**: All authorization failures should be logged for security monitoring
3. **Regular Reviews**: Periodically review access control matrix to ensure it matches business requirements
4. **Testing**: Implement automated tests for access control to prevent regressions
5. **Documentation**: Keep access control matrix updated as new features are added

## Notes

- All authorization checks are enforced at the route level
- Department-based access control is the primary authorization mechanism
- Horizontal privilege escalation is prevented through user_id validation
- Vertical privilege escalation is prevented through department checks
- All fixes maintain backward compatibility and existing functionality

