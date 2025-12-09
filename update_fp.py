import sys
sys.stdout = open('fp_update_log.txt', 'w')
sys.stderr = sys.stdout

from app import app, db, User, EmployeeData

with app.app_context():
    print("Starting fingerprint update...")
    
    hr = User.query.filter_by(username='hr_user').first()
    if not hr:
        hr = User.query.filter_by(department='HR').first()
    
    if not hr:
        print("ERROR: HR Manager not found!")
        sys.exit(1)
    
    print(f"Found HR Manager: {hr.username} (ID: {hr.id})")
    
    emp = EmployeeData.query.filter_by(user_id=hr.id).first()
    if not emp:
        print("Creating EmployeeData record...")
        emp = EmployeeData(user_id=hr.id)
        db.session.add(emp)
    
    old_fp = emp.browser_fingerprint
    print(f"Old fingerprint: {old_fp}")
    
    new_fp = "6fc7c55b2ee9afd4dd8e9454b3a93ca6"
    print(f"Setting new fingerprint: {new_fp}")
    emp.browser_fingerprint = new_fp
    
    db.session.commit()
    print("Committed to database")
    
    # Reload and verify
    db.session.refresh(emp)
    stored = emp.browser_fingerprint
    print(f"Stored fingerprint: {stored}")
    print(f"Match: {stored == new_fp}")
    
    print("Done!")

sys.stdout.close()

