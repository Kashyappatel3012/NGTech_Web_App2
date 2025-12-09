"""
Quick test to check and update HR Manager fingerprint
"""
import sys
from app import app, db, User, EmployeeData

with app.app_context():
    # Find HR Manager
    hr_manager = User.query.filter_by(username='hr_user').first()
    if not hr_manager:
        hr_manager = User.query.filter_by(department='HR').first()
    
    if not hr_manager:
        print("❌ HR Manager not found!")
        sys.exit(1)
    
    print(f"✅ Found HR Manager: {hr_manager.username} (ID: {hr_manager.id})")
    
    # Get employee data
    employee_data = EmployeeData.query.filter_by(user_id=hr_manager.id).first()
    if not employee_data:
        print("❌ Employee data not found!")
        sys.exit(1)
    
    print(f"\n📋 Current stored fingerprint:")
    if employee_data.browser_fingerprint:
        fp = employee_data.browser_fingerprint
        print(f"   Value: {fp}")
        print(f"   Length: {len(fp)} characters")
        
        # Check if it's encrypted (long base64) or plain text (32 char MD5)
        if len(fp) > 50:
            print(f"   Type: Likely encrypted (long string)")
        elif len(fp) == 32:
            print(f"   Type: Plain text MD5")
            print(f"   ✅ This should work with the validation code!")
        else:
            print(f"   Type: Unknown format")
    else:
        print("   No fingerprint stored!")
    
    # Update to plain text
    new_fp = "6fc7c55b2ee9afd4dd8e9454b3a93ca6"
    print(f"\n📝 Updating to: {new_fp}")
    employee_data.browser_fingerprint = new_fp
    db.session.commit()
    
    print(f"✅ Updated successfully!")
    
    # Verify
    verify = EmployeeData.query.filter_by(user_id=hr_manager.id).first()
    if verify and verify.browser_fingerprint == new_fp:
        print(f"✅ Verification passed! Fingerprint matches.")
    else:
        print(f"❌ Verification failed!")

