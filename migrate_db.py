"""Database migration script to rename mac_address to browser_fingerprint"""
from app import app, db

with app.app_context():
    try:
        # Try to rename mac_address column to browser_fingerprint if it exists
        db.session.execute(db.text('ALTER TABLE employee_data RENAME COLUMN mac_address TO browser_fingerprint'))
        db.session.commit()
        print("SUCCESS: Renamed mac_address column to browser_fingerprint")
    except Exception as e:
        error_msg = str(e).lower()
        if 'no such column' in error_msg or 'mac_address' not in error_msg:
            # Column doesn't exist, try to add browser_fingerprint
            try:
                db.session.execute(db.text('ALTER TABLE employee_data ADD COLUMN browser_fingerprint VARCHAR(50)'))
                db.session.commit()
                print("SUCCESS: Added browser_fingerprint column to employee_data table")
            except Exception as e2:
                error_msg2 = str(e2).lower()
                if 'duplicate column' in error_msg2 or 'already exists' in error_msg2:
                    print("SUCCESS: browser_fingerprint column already exists - no migration needed")
                elif 'no such table' in error_msg2:
                    print("WARNING: employee_data table doesn't exist yet - it will be created on first run")
                else:
                    print(f"ERROR: {e2}")
                    db.session.rollback()
        elif 'no such table' in error_msg:
            print("WARNING: employee_data table doesn't exist yet - it will be created on first run")
        else:
            print(f"ERROR: {e}")
            db.session.rollback()

