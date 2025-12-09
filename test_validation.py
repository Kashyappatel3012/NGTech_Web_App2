"""Test if fingerprint validation works"""
import sys
sys.stdout = open('validation_test.txt', 'w')
sys.stderr = sys.stdout

from app import app, db, User, EmployeeData
from app import validate_browser_fingerprint

with app.app_context():
    print("Testing fingerprint validation...")
    
    # Find HR Manager
    hr = User.query.filter_by(username='hr_user').first()
    if not hr:
        hr = User.query.filter_by(department='HR').first()
    
    if not hr:
        print("ERROR: HR Manager not found!")
        sys.exit(1)
    
    print(f"HR Manager: {hr.username} (ID: {hr.id})")
    
    # Get stored fingerprint
    emp = EmployeeData.query.filter_by(user_id=hr.id).first()
    if emp and emp.browser_fingerprint:
        stored = emp.browser_fingerprint
        print(f"Stored fingerprint: {stored}")
        print(f"Length: {len(stored)}")
    else:
        print("ERROR: No fingerprint stored!")
        sys.exit(1)
    
    # Test validation with correct fingerprint
    test_fp = "6fc7c55b2ee9afd4dd8e9454b3a93ca6"
    print(f"\nTesting with fingerprint: {test_fp}")
    
    is_valid, user_found = validate_browser_fingerprint(test_fp, hr)
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

