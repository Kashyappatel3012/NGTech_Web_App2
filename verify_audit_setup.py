"""
Verify Audit User fingerprint setup is correct
"""
import sys
from app import app, db, User, EmployeeData
from app import validate_browser_fingerprint

def verify_audit_setup():
    with app.app_context():
        print("=" * 60)
        print("Audit User Fingerprint Verification")
        print("=" * 60)
        
        # Find audit user
        audit_user = User.query.filter_by(username='audit_user').first()
        if not audit_user:
            audit_user = User.query.filter_by(department='Audit').first()
        
        if not audit_user:
            print("❌ ERROR: Audit user not found!")
            return False
        
        print(f"✅ Found Audit User:")
        print(f"   - Username: {audit_user.username}")
        print(f"   - Name: {audit_user.employee_name}")
        print(f"   - Department: {audit_user.department}")
        print(f"   - User ID: {audit_user.id}")
        
        # Get employee data
        emp = EmployeeData.query.filter_by(user_id=audit_user.id).first()
        if not emp:
            print("❌ ERROR: Employee data not found!")
            return False
        
        # Check fingerprint
        expected_fp = "d492bdacd9c3dc403c9d792a1e70feb0"
        stored_fp = emp.browser_fingerprint
        
        print(f"\n📋 Fingerprint Information:")
        print(f"   - Expected: {expected_fp}")
        print(f"   - Stored:   {stored_fp if stored_fp else 'None'}")
        print(f"   - Length:   {len(stored_fp) if stored_fp else 0} characters")
        print(f"   - Format:   {'MD5 (32 chars)' if stored_fp and len(stored_fp) == 32 else 'Unknown' if stored_fp else 'None'}")
        
        if not stored_fp:
            print("\n❌ ERROR: No fingerprint stored in database!")
            return False
        
        if stored_fp != expected_fp:
            print(f"\n⚠️  WARNING: Stored fingerprint doesn't match expected value!")
            print(f"   Update needed!")
            return False
        
        print(f"\n✅ Fingerprint matches expected value!")
        
        # Test validation with user
        print(f"\n🧪 Testing Validation (with user parameter):")
        is_valid, user_found = validate_browser_fingerprint(expected_fp, audit_user)
        if is_valid and user_found and user_found.username == 'audit_user':
            print(f"   ✅ Validation PASSED")
            print(f"   ✅ User found: {user_found.username}")
        else:
            print(f"   ❌ Validation FAILED")
            print(f"   is_valid: {is_valid}")
            print(f"   user_found: {user_found.username if user_found else 'None'}")
            return False
        
        # Test validation without user (like during login)
        print(f"\n🧪 Testing Validation (without user parameter):")
        is_valid2, user_found2 = validate_browser_fingerprint(expected_fp, None)
        if is_valid2 and user_found2 and user_found2.username == 'audit_user':
            print(f"   ✅ Validation PASSED")
            print(f"   ✅ User found: {user_found2.username}")
        else:
            print(f"   ❌ Validation FAILED")
            print(f"   is_valid: {is_valid2}")
            print(f"   user_found: {user_found2.username if user_found2 else 'None'}")
            return False
        
        print(f"\n" + "=" * 60)
        print("✅ ALL TESTS PASSED!")
        print("=" * 60)
        print(f"\nAudit User fingerprint is correctly configured.")
        print(f"Fingerprint: {expected_fp}")
        print(f"Login should work correctly with this fingerprint.")
        
        return True

if __name__ == '__main__':
    success = verify_audit_setup()
    sys.exit(0 if success else 1)

