"""
Script to set up production database on Render.com
Run this in Render Shell after database is created
"""
import sys
import io

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from app import app, db, User, EmployeeData, UserStatus
from werkzeug.security import generate_password_hash
from datetime import datetime

def setup_production_database():
    """Set up initial users and data in production database"""
    with app.app_context():
        print("=" * 70)
        print("Production Database Setup")
        print("=" * 70)
        
        # Check database connection
        try:
            db.session.execute(db.text("SELECT 1"))
            print("\n[OK] Database connection successful!")
        except Exception as e:
            print(f"\n[ERROR] Database connection failed: {e}")
            return False
        
        # Check existing tables
        from sqlalchemy import inspect
        inspector = inspect(db.engine)
        tables = inspector.get_table_names()
        print(f"\n[INFO] Found {len(tables)} tables in database:")
        for table in tables:
            print(f"  - {table}")
        
        # Create tables if they don't exist
        if 'user' not in tables:
            print("\n[INFO] Creating database tables...")
            try:
                db.create_all()
                db.session.commit()
                print("[OK] Tables created successfully!")
            except Exception as e:
                print(f"[ERROR] Failed to create tables: {e}")
                return False
        
        # Check existing users
        existing_users = User.query.all()
        print(f"\n[INFO] Found {len(existing_users)} existing user(s)")
        
        # Setup HR Manager
        hr_user = User.query.filter_by(username='hr_user').first()
        if not hr_user:
            print("\n[INFO] Creating HR Manager user...")
            try:
                hr_user = User(
                    username='hr_user',
                    employee_name='HR Manager',
                    password=generate_password_hash('hr123'),  # Change this password!
                    email='pubglover30120101@gmail.com',
                    department='HR',
                    created_at=datetime.now()
                )
                db.session.add(hr_user)
                db.session.flush()
                
                # Create EmployeeData
                emp_data = EmployeeData(
                    user_id=hr_user.id,
                    browser_fingerprint='396520d70ea1f79dd21caffd85085795',  # Update with your fingerprint
                    position='HR Manager',
                    created_at=datetime.now()
                )
                db.session.add(emp_data)
                
                # Create UserStatus
                user_status = UserStatus(
                    user_id=hr_user.id,
                    is_active=True,
                    failed_attempts=0
                )
                db.session.add(user_status)
                
                db.session.commit()
                print("[OK] HR Manager created successfully!")
                print(f"  - Username: hr_user")
                print(f"  - Password: hr123 (CHANGE THIS!)")
                print(f"  - Fingerprint: 396520d70ea1f79dd21caffd85085795")
            except Exception as e:
                db.session.rollback()
                print(f"[ERROR] Failed to create HR Manager: {e}")
                return False
        else:
            print("\n[INFO] HR Manager already exists")
            # Update fingerprint if needed
            emp_data = EmployeeData.query.filter_by(user_id=hr_user.id).first()
            if emp_data:
                old_fp = emp_data.browser_fingerprint
                emp_data.browser_fingerprint = '396520d70ea1f79dd21caffd85085795'
                db.session.commit()
                if old_fp != emp_data.browser_fingerprint:
                    print(f"[OK] Updated fingerprint: {old_fp} -> {emp_data.browser_fingerprint}")
                else:
                    print(f"[INFO] Fingerprint already correct: {emp_data.browser_fingerprint}")
        
        # Setup Admin User (Kashyap Patel)
        admin_user = User.query.filter_by(username='kashyap.patel').first()
        if not admin_user:
            print("\n[INFO] Creating Admin user (Kashyap Patel)...")
            try:
                admin_user = User(
                    username='kashyap.patel',
                    employee_name='Kashyap Patel',
                    password=generate_password_hash('Admin@2024'),  # Change this password!
                    email='patelkashyap3012@gmail.com',
                    department='Admin',
                    created_at=datetime.now()
                )
                db.session.add(admin_user)
                db.session.flush()
                
                # Create EmployeeData
                emp_data = EmployeeData(
                    user_id=admin_user.id,
                    browser_fingerprint='24769342a752806361471a8e6db5f78d',  # Update with your fingerprint
                    position='Admin',
                    created_at=datetime.now()
                )
                db.session.add(emp_data)
                
                # Create UserStatus
                user_status = UserStatus(
                    user_id=admin_user.id,
                    is_active=True,
                    failed_attempts=0
                )
                db.session.add(user_status)
                
                db.session.commit()
                print("[OK] Admin user created successfully!")
                print(f"  - Username: kashyap.patel")
                print(f"  - Password: Admin@2024 (CHANGE THIS!)")
                print(f"  - Fingerprint: 24769342a752806361471a8e6db5f78d")
            except Exception as e:
                db.session.rollback()
                print(f"[ERROR] Failed to create Admin user: {e}")
        
        # Final summary
        print("\n" + "=" * 70)
        print("Setup Summary")
        print("=" * 70)
        all_users = User.query.all()
        print(f"\nTotal users in database: {len(all_users)}")
        for user in all_users:
            emp = EmployeeData.query.filter_by(user_id=user.id).first()
            fp = emp.browser_fingerprint if emp else "None"
            print(f"  - {user.username} ({user.department})")
            print(f"    Fingerprint: {fp}")
        
        print("\n" + "=" * 70)
        print("[OK] Database setup completed!")
        print("=" * 70)
        print("\n⚠️  IMPORTANT: Change default passwords after first login!")
        
        return True

if __name__ == '__main__':
    try:
        success = setup_production_database()
        if not success:
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n[ERROR] Operation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

