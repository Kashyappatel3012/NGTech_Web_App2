"""Direct fix for HR Manager fingerprint - store unencrypted"""
from app import app, db, User, EmployeeData

def fix_fingerprint():
    with app.app_context():
        # Find HR Manager
        hr = User.query.filter_by(username='hr_user').first() or User.query.filter_by(department='HR').first()
        
        if not hr:
            print("HR Manager not found!")
            return
        
        print(f"Found: {hr.username}")
        
        # Get or create employee data
        emp = EmployeeData.query.filter_by(user_id=hr.id).first()
        if not emp:
            emp = EmployeeData(user_id=hr.id)
            db.session.add(emp)
        
        # Store unencrypted (plain text) for compatibility
        new_fp = "6fc7c55b2ee9afd4dd8e9454b3a93ca6"
        emp.browser_fingerprint = new_fp
        
        db.session.commit()
        
        # Verify
        check = EmployeeData.query.filter_by(user_id=hr.id).first()
        if check.browser_fingerprint == new_fp:
            print(f"SUCCESS! Fingerprint set to: {new_fp}")
            print(f"Value in DB: {check.browser_fingerprint}")
        else:
            print("FAILED!")

if __name__ == '__main__':
    fix_fingerprint()

