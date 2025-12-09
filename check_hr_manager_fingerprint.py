"""
Script to check HR Manager browser fingerprint in database
"""
import os
import sys
from app import app, db, User, EmployeeData
from encryption_utils import get_encryption_manager

def check_hr_manager_fingerprint():
    """Check HR Manager's browser fingerprint in database"""
    with app.app_context():
        # Find HR Manager user - try different usernames
        hr_manager = User.query.filter_by(username='hr_user').first()
        if not hr_manager:
            # Try finding by department
            hr_manager = User.query.filter_by(department='HR').first()
        if not hr_manager:
            # List all users to see what we have
            all_users = User.query.all()
            print(f"❌ HR Manager user not found!")
            print(f"Available users: {[u.username for u in all_users]}")
            return
        
        print(f"✅ Found HR Manager user: {hr_manager.username} (ID: {hr_manager.id})")
        
        # Get employee data
        employee_data = EmployeeData.query.filter_by(user_id=hr_manager.id).first()
        if not employee_data:
            print("❌ Employee data not found for HR Manager!")
            return
        
        print(f"✅ Found employee data for HR Manager")
        
        # Check browser fingerprint
        stored_fingerprint_encrypted = employee_data.browser_fingerprint
        if not stored_fingerprint_encrypted:
            print("❌ No browser fingerprint stored in database!")
            return
        
        print(f"\n📋 Stored Fingerprint (Encrypted): {stored_fingerprint_encrypted[:50]}...")
        
        # Try to decrypt
        enc_manager = get_encryption_manager()
        try:
            stored_fingerprint = enc_manager.decrypt(stored_fingerprint_encrypted)
            print(f"✅ Decrypted Fingerprint: {stored_fingerprint}")
            print(f"   Length: {len(stored_fingerprint)} characters")
            
            # Check if it's MD5 (32 chars) or SHA-256 (64 chars)
            if len(stored_fingerprint) == 32:
                print("   Type: MD5 (32 characters)")
                print(f"   ⚠️  WARNING: Browser now generates SHA-256 (64 characters), so they won't match!")
            elif len(stored_fingerprint) == 64:
                print("   Type: SHA-256 (64 characters)")
                print(f"   ✅ Matches current browser fingerprint format")
            else:
                print(f"   Type: Unknown format")
            
            # Check if it matches expected value
            expected_md5 = "485fdd03538a8c43174d9c57819143b8"
            if stored_fingerprint == expected_md5:
                print(f"\n✅ Stored fingerprint matches expected MD5: {expected_md5}")
                print(f"   ⚠️  However, browser now generates SHA-256, so login will fail!")
            else:
                print(f"\n❌ Stored fingerprint does NOT match expected MD5: {expected_md5}")
                print(f"   Expected: {expected_md5}")
                print(f"   Stored:   {stored_fingerprint}")
        
        except Exception as e:
            print(f"❌ Error decrypting fingerprint: {e}")
            print(f"   Trying direct comparison (might be unencrypted)...")
            if stored_fingerprint_encrypted == "485fdd03538a8c43174d9c57819143b8":
                print(f"   ✅ Matches expected MD5 (unencrypted)")
            else:
                print(f"   ❌ Does not match expected MD5")

if __name__ == '__main__':
    check_hr_manager_fingerprint()

