"""
Script to update HR Manager browser fingerprint from MD5 to SHA-256
This is needed because we changed the fingerprint algorithm from MD5 to SHA-256
"""
import os
import sys
from app import app, db, User, EmployeeData
from encryption_utils import get_encryption_manager
import hashlib

def generate_md5_fingerprint_from_components():
    """
    Generate MD5 fingerprint using the same method as before
    This is for reference - we need to know what the MD5 was to convert it
    """
    # This is just for reference - the actual fingerprint components would be
    # the same as what the browser collects
    pass

def update_hr_manager_fingerprint():
    """
    Update HR Manager's fingerprint to SHA-256 format
    Since we can't regenerate the exact SHA-256 from MD5, we need to:
    1. Check current stored fingerprint
    2. Either update it manually or provide instructions
    """
    with app.app_context():
        # Find HR Manager user
        hr_manager = User.query.filter_by(username='hr_user').first()
        if not hr_manager:
            # Try by department
            hr_manager = User.query.filter_by(department='HR').first()
        
        if not hr_manager:
            print("❌ HR Manager user not found!")
            all_users = User.query.all()
            print(f"Available users: {[u.username for u in all_users]}")
            return
        
        print(f"✅ Found HR Manager: {hr_manager.username} (ID: {hr_manager.id})")
        
        # Get employee data
        employee_data = EmployeeData.query.filter_by(user_id=hr_manager.id).first()
        if not employee_data:
            print("❌ Employee data not found!")
            return
        
        # Check current fingerprint
        stored_fp_encrypted = employee_data.browser_fingerprint
        if not stored_fp_encrypted:
            print("❌ No fingerprint stored!")
            return
        
        enc_manager = get_encryption_manager()
        try:
            stored_fp = enc_manager.decrypt(stored_fp_encrypted)
            print(f"\n📋 Current Stored Fingerprint: {stored_fp}")
            print(f"   Length: {len(stored_fp)} characters")
            
            if len(stored_fp) == 32:
                print("   Type: MD5 (32 characters) - OLD FORMAT")
                print("\n⚠️  ISSUE FOUND:")
                print("   The stored fingerprint is MD5 (32 chars), but the browser")
                print("   now generates SHA-256 (64 chars), so they won't match!")
                print("\n💡 SOLUTION:")
                print("   You need to:")
                print("   1. Open the browser that should have access")
                print("   2. Go to the login page")
                print("   3. Check the browser fingerprint displayed (it will be SHA-256, 64 chars)")
                print("   4. Update the database with that new SHA-256 fingerprint")
                print("\n   Or temporarily revert to MD5 in login.js/verify_otp.js")
                
                # Ask if user wants to update
                print("\n❓ Do you want to update the fingerprint?")
                print("   You'll need to provide the new SHA-256 fingerprint from the browser.")
                print("   Run this script with the new fingerprint as argument:")
                print("   python update_fingerprint_to_sha256.py <new_sha256_fingerprint>")
                
                if len(sys.argv) > 1:
                    new_fingerprint = sys.argv[1].strip()
                    if len(new_fingerprint) == 64:
                        # Update to new SHA-256 fingerprint
                        encrypted_new = enc_manager.encrypt(new_fingerprint)
                        employee_data.browser_fingerprint = encrypted_new
                        db.session.commit()
                        print(f"\n✅ Updated fingerprint to SHA-256: {new_fingerprint}")
                        print("   Database updated successfully!")
                    else:
                        print(f"\n❌ Invalid fingerprint length: {len(new_fingerprint)} (expected 64 for SHA-256)")
                else:
                    print("\n   Current stored MD5: " + stored_fp)
                    print("   Expected format: SHA-256 (64 hexadecimal characters)")
            
            elif len(stored_fp) == 64:
                print("   Type: SHA-256 (64 characters) - NEW FORMAT")
                print("   ✅ Fingerprint is already in SHA-256 format!")
            else:
                print(f"   Type: Unknown format ({len(stored_fp)} characters)")
        
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()

if __name__ == '__main__':
    update_hr_manager_fingerprint()

