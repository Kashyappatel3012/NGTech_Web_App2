"""Direct fix for Audit User fingerprint - store unencrypted"""
import sys
sys.stdout = open('audit_update.txt', 'w')
sys.stderr = sys.stdout

from app import app, db, User, EmployeeData

def fix_fingerprint():
    with app.app_context():
        # Find Audit User
        audit = User.query.filter_by(username='audit_user').first()
        if not audit:
            audit = User.query.filter_by(department='Audit').first()
        
        if not audit:
            print("Audit User not found!")
            return False
        
        print(f"Found: {audit.username} (ID: {audit.id})")
        
        # Get or create employee data
        emp = EmployeeData.query.filter_by(user_id=audit.id).first()
        if not emp:
            emp = EmployeeData(user_id=audit.id)
            db.session.add(emp)
        
        # Store unencrypted (plain text) for compatibility
        new_fp = "d492bdacd9c3dc403c9d792a1e70feb0"
        old_fp = emp.browser_fingerprint
        print(f"Old fingerprint: {old_fp}")
        print(f"Setting new fingerprint: {new_fp}")
        emp.browser_fingerprint = new_fp
        
        db.session.commit()
        print("Committed to database")
        
        # Reload and verify
        db.session.refresh(emp)
        stored = emp.browser_fingerprint
        print(f"Stored fingerprint: {stored}")
        print(f"Match: {stored == new_fp}")
        
        if stored == new_fp:
            print("SUCCESS! Fingerprint updated correctly.")
            return True
        else:
            print("FAILED! Fingerprint doesn't match.")
            return False

if __name__ == '__main__':
    success = fix_fingerprint()
    sys.stdout.close()
    if success:
        print("\n✅ Audit User fingerprint updated successfully!")
    else:
        print("\n❌ Failed to update Audit User fingerprint!")

