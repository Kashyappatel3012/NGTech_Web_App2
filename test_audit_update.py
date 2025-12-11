"""Test script to update and verify audit user fingerprint"""
from app import app, db, User, EmployeeData
from app import validate_browser_fingerprint

with app.app_context():
    # Find audit user
    audit = User.query.filter_by(username='audit_user').first()
    if not audit:
        audit = User.query.filter_by(department='Audit').first()
    
    if not audit:
        print("ERROR: Audit user not found")
        exit(1)
    
    print(f"Found audit user: {audit.username} (ID: {audit.id})")
    
    # Get or create employee data
    emp = EmployeeData.query.filter_by(user_id=audit.id).first()
    if not emp:
        emp = EmployeeData(user_id=audit.id)
        db.session.add(emp)
        db.session.commit()
    
    # Update fingerprint
    new_fp = "d492bdacd9c3dc403c9d792a1e70feb0"
    print(f"Updating fingerprint to: {new_fp}")
    emp.browser_fingerprint = new_fp
    db.session.commit()
    
    # Verify
    db.session.refresh(emp)
    print(f"Stored fingerprint: {emp.browser_fingerprint}")
    print(f"Match: {emp.browser_fingerprint == new_fp}")
    
    # Test validation
    is_valid, user_found = validate_browser_fingerprint(new_fp, audit)
    print(f"Validation test: valid={is_valid}, user={user_found.username if user_found else None}")
    
    if is_valid and user_found and user_found.username == 'audit_user':
        print("SUCCESS: Fingerprint updated and validated correctly!")
    else:
        print("WARNING: Validation may have issues")

