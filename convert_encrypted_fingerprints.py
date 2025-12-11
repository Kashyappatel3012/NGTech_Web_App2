"""
Script to convert encrypted browser fingerprints to unencrypted format in database.
This ensures fingerprints work consistently even if encryption key changes.
"""
import sys
from app import app, db, EmployeeData
from encryption_utils import get_encryption_manager

def convert_encrypted_fingerprints():
    """
    Convert all encrypted browser fingerprints to unencrypted format.
    This fixes the issue where fingerprints fail after logout/timeout when encryption key changes.
    """
    with app.app_context():
        print("=" * 70)
        print("Converting Encrypted Browser Fingerprints to Unencrypted Format")
        print("=" * 70)
        print()
        
        # Get all employee data with fingerprints
        all_employee_data = EmployeeData.query.filter(
            EmployeeData.browser_fingerprint.isnot(None)
        ).all()
        
        if not all_employee_data:
            print("No fingerprints found in database.")
            return True
        
        print(f"Found {len(all_employee_data)} employee record(s) with fingerprints.")
        print()
        
        # Try to get encryption manager (for decrypting old encrypted data)
        enc_manager = None
        try:
            enc_manager = get_encryption_manager()
            print("Encryption manager available - will attempt to decrypt old encrypted fingerprints.")
        except Exception as e:
            print(f"Warning: Encryption manager not available: {e}")
            print("Will treat all fingerprints as plain text.")
            print()
        
        converted_count = 0
        already_unencrypted_count = 0
        failed_count = 0
        
        for emp_data in all_employee_data:
            stored_value = emp_data.browser_fingerprint
            plaintext_fingerprint = None
            
            print(f"Processing Employee ID: {emp_data.user_id}")
            print(f"  Stored value: {stored_value[:50]}..." if len(stored_value) > 50 else f"  Stored value: {stored_value}")
            
            # Try to determine if it's encrypted or plain text
            is_encrypted = False
            if enc_manager:
                try:
                    # Try to decrypt - if it succeeds, it was encrypted
                    plaintext_fingerprint = enc_manager.decrypt(stored_value)
                    is_encrypted = True
                    print(f"  Status: Encrypted (decrypted successfully)")
                except:
                    # Decryption failed - assume it's already plain text
                    plaintext_fingerprint = stored_value
                    is_encrypted = False
                    print(f"  Status: Already unencrypted (decryption failed, assuming plain text)")
            else:
                # No encryption manager - assume it's already plain text
                plaintext_fingerprint = stored_value
                is_encrypted = False
                print(f"  Status: Treating as plain text (no encryption manager)")
            
            # Check if it needs conversion
            if is_encrypted and plaintext_fingerprint:
                # Update to unencrypted format
                try:
                    emp_data.browser_fingerprint = plaintext_fingerprint.strip()
                    converted_count += 1
                    print(f"  ✅ Converted to unencrypted: {plaintext_fingerprint}")
                except Exception as e:
                    failed_count += 1
                    print(f"  ❌ Failed to convert: {e}")
            elif not is_encrypted:
                already_unencrypted_count += 1
                print(f"  ✓ Already unencrypted: {plaintext_fingerprint}")
            
            print()
        
        # Commit all changes
        if converted_count > 0:
            try:
                db.session.commit()
                print("=" * 70)
                print("Conversion Summary:")
                print(f"  ✅ Converted: {converted_count} fingerprint(s)")
                print(f"  ✓ Already unencrypted: {already_unencrypted_count} fingerprint(s)")
                if failed_count > 0:
                    print(f"  ❌ Failed: {failed_count} fingerprint(s)")
                print("=" * 70)
                print()
                print("✅ SUCCESS! All fingerprints are now stored as unencrypted in the database.")
                print("   Fingerprints will work consistently even if encryption key changes.")
                return True
            except Exception as e:
                db.session.rollback()
                print("=" * 70)
                print(f"❌ ERROR: Failed to commit changes: {e}")
                print("=" * 70)
                return False
        else:
            print("=" * 70)
            print("No conversions needed - all fingerprints are already unencrypted.")
            print("=" * 70)
            return True

if __name__ == '__main__':
    print()
    print("Starting fingerprint conversion...")
    print()
    
    success = convert_encrypted_fingerprints()
    
    if success:
        print()
        print("✅ Script completed successfully!")
        sys.exit(0)
    else:
        print()
        print("❌ Script failed!")
        sys.exit(1)

