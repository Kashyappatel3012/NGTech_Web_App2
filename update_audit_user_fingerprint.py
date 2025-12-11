"""
Script to update Audit User browser fingerprint in database
"""
import os
import sys
from app import app, db, User, EmployeeData
from encryption_utils import get_encryption_manager

def update_audit_user_fingerprint(new_fingerprint):
    """
    Update Audit User's browser fingerprint in database
    
    Args:
        new_fingerprint: The new browser fingerprint to set (e.g., 'd492bdacd9c3dc403c9d792a1e70feb0')
    """
    with app.app_context():
        # Find Audit User - try different methods
        audit_user = None
        
        # Try finding by username 'audit_user'
        audit_user = User.query.filter_by(username='audit_user').first()
        
        # If not found, try by department
        if not audit_user:
            audit_user = User.query.filter_by(department='Audit').first()
        
        # If still not found, list all users
        if not audit_user:
            all_users = User.query.all()
            print(f"❌ Audit User not found!")
            print(f"Available users:")
            for user in all_users:
                print(f"  - Username: {user.username}, Department: {user.department}, Name: {user.employee_name}")
            return False
        
        print(f"✅ Found Audit User:")
        print(f"   - Username: {audit_user.username}")
        print(f"   - Name: {audit_user.employee_name}")
        print(f"   - Department: {audit_user.department}")
        print(f"   - User ID: {audit_user.id}")
        
        # Get employee data
        employee_data = EmployeeData.query.filter_by(user_id=audit_user.id).first()
        if not employee_data:
            print("❌ Employee data not found for Audit User!")
            print("   Creating EmployeeData record...")
            employee_data = EmployeeData(user_id=audit_user.id)
            db.session.add(employee_data)
        
        # Show current fingerprint if it exists
        enc_manager = get_encryption_manager()
        if employee_data.browser_fingerprint:
            try:
                current_fingerprint = enc_manager.decrypt(employee_data.browser_fingerprint)
                print(f"\n📋 Current Fingerprint (Decrypted): {current_fingerprint}")
                print(f"   Length: {len(current_fingerprint)} characters")
            except Exception as e:
                print(f"\n📋 Current Fingerprint (Encrypted): {employee_data.browser_fingerprint[:50]}...")
                print(f"   (Could not decrypt - may be in old format or corrupted)")
        else:
            print("\n📋 No current fingerprint stored in database")
        
        # Validate new fingerprint format
        new_fingerprint = new_fingerprint.strip()
        if not new_fingerprint:
            print("❌ Error: New fingerprint cannot be empty!")
            return False
        
        print(f"\n📝 New Fingerprint: {new_fingerprint}")
        print(f"   Length: {len(new_fingerprint)} characters")
        
        # Check if it's MD5 (32 chars) or SHA-256 (64 chars)
        if len(new_fingerprint) == 32:
            print("   Type: MD5 (32 characters)")
        elif len(new_fingerprint) == 64:
            print("   Type: SHA-256 (64 characters)")
        else:
            print(f"   Type: Unknown format ({len(new_fingerprint)} characters)")
            print("   ⚠️  Warning: Expected MD5 (32 chars) or SHA-256 (64 chars)")
        
        # Update the fingerprint
        # NOTE: Storing unencrypted for compatibility with existing system
        # The validation code handles both encrypted and unencrypted fingerprints
        # If encryption key changes between script runs and app runs, encrypted data won't decrypt
        try:
            # Check if we should encrypt or store plain text
            # For compatibility and to avoid encryption key mismatch issues, store unencrypted
            # The validation code will handle both formats
            store_encrypted = os.environ.get('STORE_FINGERPRINT_ENCRYPTED', 'false').lower() == 'true'
            
            if store_encrypted:
                encrypted_fingerprint = enc_manager.encrypt(new_fingerprint)
                employee_data.browser_fingerprint = encrypted_fingerprint
                print(f"   - Stored as: Encrypted")
            else:
                # Store unencrypted for compatibility (matches old format)
                employee_data.browser_fingerprint = new_fingerprint
                print(f"   - Stored as: Unencrypted (for compatibility)")
            
            # Commit to database
            db.session.commit()
            
            print(f"\n✅ SUCCESS! Fingerprint updated in database")
            print(f"   - User: {audit_user.username} ({audit_user.employee_name})")
            print(f"   - New Fingerprint: {new_fingerprint}")
            print(f"   - Stored successfully")
            
            # Verify the update
            verify_employee_data = EmployeeData.query.filter_by(user_id=audit_user.id).first()
            if verify_employee_data and verify_employee_data.browser_fingerprint:
                stored_value = verify_employee_data.browser_fingerprint
                
                # Try to decrypt first (in case it's encrypted)
                try:
                    verify_fingerprint = enc_manager.decrypt(stored_value)
                except:
                    # If decryption fails, assume it's unencrypted
                    verify_fingerprint = stored_value
                
                if verify_fingerprint == new_fingerprint:
                    print(f"\n✅ Verification: Fingerprint matches!")
                    print(f"   Stored value: {stored_value[:50]}..." if len(stored_value) > 50 else f"   Stored value: {stored_value}")
                else:
                    print(f"\n⚠️  Warning: Verification failed - fingerprint doesn't match!")
                    print(f"   Expected: {new_fingerprint}")
                    print(f"   Got: {verify_fingerprint}")
            
            return True
            
        except Exception as e:
            db.session.rollback()
            print(f"\n❌ ERROR: Failed to update fingerprint: {e}")
            import traceback
            traceback.print_exc()
            return False

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python update_audit_user_fingerprint.py <new_fingerprint>")
        print("Example: python update_audit_user_fingerprint.py d492bdacd9c3dc403c9d792a1e70feb0")
        sys.exit(1)
    
    new_fingerprint = sys.argv[1]
    success = update_audit_user_fingerprint(new_fingerprint)
    
    if success:
        print("\n✅ Script completed successfully!")
        sys.exit(0)
    else:
        print("\n❌ Script failed!")
        sys.exit(1)

