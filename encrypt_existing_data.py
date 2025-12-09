"""Migration script to encrypt existing browser fingerprints in database"""
from app import app, db
from app import User, EmployeeData
from encryption_utils import get_encryption_manager

with app.app_context():
    enc_manager = get_encryption_manager()
    
    print("Starting encryption of existing browser fingerprints...")
    
    # Get all employee data with browser fingerprints
    all_employee_data = EmployeeData.query.filter(EmployeeData.browser_fingerprint.isnot(None)).all()
    
    encrypted_count = 0
    skipped_count = 0
    error_count = 0
    
    for emp_data in all_employee_data:
        if not emp_data.browser_fingerprint:
            continue
        
        # Check if already encrypted (encrypted data is longer and base64-like)
        # Encrypted data will be much longer than 32 characters (MD5 hash)
        if len(emp_data.browser_fingerprint) > 50:
            # Likely already encrypted, skip
            skipped_count += 1
            continue
        
        try:
            # Encrypt the fingerprint
            encrypted_fp = enc_manager.encrypt(emp_data.browser_fingerprint)
            emp_data.browser_fingerprint = encrypted_fp
            encrypted_count += 1
            print(f"Encrypted fingerprint for user_id: {emp_data.user_id}")
        except Exception as e:
            print(f"Error encrypting fingerprint for user_id {emp_data.user_id}: {e}")
            error_count += 1
    
    try:
        db.session.commit()
        print(f"\n✅ Encryption complete!")
        print(f"   - Encrypted: {encrypted_count}")
        print(f"   - Already encrypted (skipped): {skipped_count}")
        print(f"   - Errors: {error_count}")
    except Exception as e:
        db.session.rollback()
        print(f"\n❌ Error committing changes: {e}")

