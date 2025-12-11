"""Verify Audit User fingerprint was updated correctly"""
import sys
sys.stdout = open('audit_fp_verify.txt', 'w')
sys.stderr = sys.stdout

from app import app, db, User, EmployeeData
from app import validate_browser_fingerprint

with app.app_context():
    print("Verifying Audit User fingerprint...")
    
    # Find Audit User
    audit_user = User.query.filter_by(username='audit_user').first()
    if not audit_user:
        audit_user = User.query.filter_by(department='Audit').first()
    
    if not audit_user:
        print("ERROR: Audit User not found!")
        sys.exit(1)
    
    print(f"Audit User: {audit_user.username} (ID: {audit_user.id})")
    
    # Get stored fingerprint
    emp = EmployeeData.query.filter_by(user_id=audit_user.id).first()
    if emp and emp.browser_fingerprint:
        stored = emp.browser_fingerprint
        print(f"Stored fingerprint: {stored}")
        print(f"Length: {len(stored)}")
    else:
        print("ERROR: No fingerprint stored!")
        sys.exit(1)
    
    # Test validation with correct fingerprint
    test_fp = "d492bdacd9c3dc403c9d792a1e70feb0"
    print(f"\nTesting with fingerprint: {test_fp}")
    
    is_valid, user_found = validate_browser_fingerprint(test_fp, audit_user)
    print(f"Validation result: is_valid={is_valid}, user={user_found.username if user_found else None}")
    
    if is_valid:
        print("✅ VALIDATION PASSED! Fingerprint will work for login.")
    else:
        print("❌ VALIDATION FAILED! Check the validation code.")
    
    # Test validation without user (like during initial login)
    print(f"\nTesting without user parameter...")
    is_valid2, user_found2 = validate_browser_fingerprint(test_fp, None)
    print(f"Validation result: is_valid={is_valid2}, user={user_found2.username if user_found2 else None}")
    
    if is_valid2:
        print("✅ VALIDATION PASSED!")
    else:
        print("❌ VALIDATION FAILED!")

sys.stdout.close()

