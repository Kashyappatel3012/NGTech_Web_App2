"""
Script to check production database status
Run this in Render Shell to verify database setup
"""
import sys
import io

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from app import app, db, User, EmployeeData, UserStatus
from sqlalchemy import inspect

def check_database():
    """Check database connection and status"""
    with app.app_context():
        print("=" * 70)
        print("Production Database Status Check")
        print("=" * 70)
        
        # Check connection
        try:
            db.session.execute(db.text("SELECT 1"))
            print("\n[OK] Database connection: SUCCESS")
        except Exception as e:
            print(f"\n[ERROR] Database connection: FAILED")
            print(f"  Error: {e}")
            return False
        
        # Check tables
        inspector = inspect(db.engine)
        tables = inspector.get_table_names()
        print(f"\n[INFO] Database tables: {len(tables)}")
        for table in tables:
            print(f"  - {table}")
        
        # Check users
        try:
            users = User.query.all()
            print(f"\n[INFO] Total users: {len(users)}")
            
            if len(users) == 0:
                print("\n[WARNING] No users found in database!")
                print("  Run setup_production_db.py to create initial users")
            else:
                print("\n[INFO] Users in database:")
                for user in users:
                    emp = EmployeeData.query.filter_by(user_id=user.id).first()
                    fp = emp.browser_fingerprint if emp else "None"
                    status = UserStatus.query.filter_by(user_id=user.id).first()
                    is_active = status.is_active if status else False
                    
                    print(f"\n  User: {user.username}")
                    print(f"    Name: {user.employee_name}")
                    print(f"    Department: {user.department}")
                    print(f"    Email: {user.email}")
                    print(f"    Active: {is_active}")
                    print(f"    Fingerprint: {fp}")
                    if fp and len(fp) == 32:
                        print(f"    Fingerprint Status: [OK] Valid MD5 format")
                    elif fp:
                        print(f"    Fingerprint Status: [WARNING] Length: {len(fp)} (expected 32)")
                    else:
                        print(f"    Fingerprint Status: [ERROR] Not set")
        except Exception as e:
            print(f"\n[ERROR] Failed to query users: {e}")
            return False
        
        print("\n" + "=" * 70)
        print("[OK] Database check completed!")
        print("=" * 70)
        
        return True

if __name__ == '__main__':
    try:
        success = check_database()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

