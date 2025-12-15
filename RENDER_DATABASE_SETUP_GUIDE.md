# Complete Render.com Database Setup Guide

## Overview
This guide will help you set up your PostgreSQL database on Render.com from scratch, including:
1. Creating the database
2. Connecting your web service to the database
3. Running migrations to create tables
4. Adding initial users and data

---

## Step 1: Create PostgreSQL Database on Render.com

### 1.1 Go to Render Dashboard
- Visit: https://dashboard.render.com
- Sign in to your account

### 1.2 Create New PostgreSQL Database
1. Click **"New +"** button (top right)
2. Select **"PostgreSQL"**
3. Fill in the details:
   - **Name:** `ngtech-db` (or any name you prefer)
   - **Database:** `ngtech_db` (or any name)
   - **User:** (auto-generated, or choose your own)
   - **Region:** Choose closest to you (e.g., `Oregon (US West)`)
   - **PostgreSQL Version:** Latest (usually 15 or 16)
   - **Plan:** Free (or paid for better performance)
4. Click **"Create Database"**

### 1.3 Note Your Database Connection Details
After creation, Render will show:
- **Internal Database URL** (for use within Render)
- **External Database URL** (for external tools like pgAdmin, DBeaver)
- **Host, Port, Database, User, Password**

**Important:** Save these details securely!

---

## Step 2: Connect Web Service to Database

### 2.1 Go to Your Web Service
1. In Render Dashboard, click on your **Web Service** (not the database)
2. Go to **"Environment"** tab

### 2.2 Add Database Connection
Render automatically sets `DATABASE_URL` when you link databases, but let's verify:

**Option A: Automatic Linking (Recommended)**
1. In your Web Service → **"Settings"** tab
2. Scroll to **"Linked Resources"**
3. Click **"Link Resource"**
4. Select your PostgreSQL database
5. Render will automatically set `DATABASE_URL`

**Option B: Manual Setup**
If automatic linking doesn't work:
1. Go to your **PostgreSQL database** → **"Connections"** tab
2. Copy the **"Internal Database URL"** (looks like: `postgresql://user:password@host:port/database`)
3. Go to your **Web Service** → **"Environment"** tab
4. Add environment variable:
   - **Key:** `DATABASE_URL`
   - **Value:** Paste the Internal Database URL
5. Click **"Save Changes"**

---

## Step 3: Verify Database Connection

### 3.1 Check Environment Variables
In your Web Service → **"Environment"** tab, you should see:
```
DATABASE_URL=postgresql://user:password@host:port/database
```

### 3.2 Check Logs
1. Go to your Web Service → **"Logs"** tab
2. Look for: `[OK] Using PostgreSQL database (Production mode)`
3. If you see this, connection is working!

---

## Step 4: Initialize Database Tables

Your Flask app automatically creates tables on first run, but let's ensure it works:

### 4.1 Automatic Table Creation
When your app starts, it calls `initialize_database()` which:
- Creates all tables (User, EmployeeData, LoginActivity, etc.)
- Runs migrations
- Creates initial dummy users (if database is empty)

**To trigger this:**
1. Deploy your app (or it will auto-deploy if connected to GitHub)
2. Check logs for: `Database initialized successfully`

### 4.2 Manual Table Creation (If Needed)
If tables aren't created automatically, you can run migrations manually:

**Option A: Using Render Shell**
1. Go to your Web Service → **"Shell"** tab
2. Run:
   ```bash
   python
   ```
3. Then in Python:
   ```python
   from app import app, db
   with app.app_context():
       db.create_all()
       print("Tables created!")
   ```

**Option B: Using Render CLI**
```bash
# Install Render CLI
npm install -g render-cli

# Login
render login

# Link to your service
render link

# Run migrations
render exec python -c "from app import app, db; app.app_context().push(); db.create_all()"
```

---

## Step 5: Add Initial Users and Data

### 5.1 Check Current Database State
First, let's see what's in the database:

**Using Render Shell:**
```bash
# In Render Shell
python
```

```python
from app import app, db, User, EmployeeData
with app.app_context():
    users = User.query.all()
    print(f"Total users: {len(users)}")
    for user in users:
        print(f"- {user.username} ({user.department})")
```

### 5.2 Add HR Manager User
If database is empty, you need to add users. Here are your options:

**Option A: Using Python Script (Recommended)**
1. Create a script `setup_production_db.py`:
   ```python
   from app import app, db, User, EmployeeData, UserStatus
   from werkzeug.security import generate_password_hash
   from datetime import datetime
   
   with app.app_context():
       # Check if HR Manager exists
       hr_user = User.query.filter_by(username='hr_user').first()
       
       if not hr_user:
           # Create HR Manager
           hr_user = User(
               username='hr_user',
               employee_name='HR Manager',
               password=generate_password_hash('hr123'),  # Change this!
               email='pubglover30120101@gmail.com',
               department='HR',
               created_at=datetime.now()
           )
           db.session.add(hr_user)
           db.session.flush()
           
           # Create EmployeeData
           emp_data = EmployeeData(
               user_id=hr_user.id,
               browser_fingerprint='396520d70ea1f79dd21caffd85085795',  # Your fingerprint
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
           print("✅ HR Manager created!")
       else:
           print("HR Manager already exists")
   ```

2. Run it in Render Shell:
   ```bash
   python setup_production_db.py
   ```

**Option B: Using SQL Directly**
1. Go to your PostgreSQL database → **"Connect"** tab
2. Use **"psql"** or **"pgAdmin"** connection string
3. Connect using external URL
4. Run SQL:
   ```sql
   -- Insert HR Manager user
   INSERT INTO "user" (username, employee_name, password, email, department, created_at)
   VALUES (
       'hr_user',
       'HR Manager',
       'pbkdf2:sha256:600000$...',  -- You need to hash password first
       'pubglover30120101@gmail.com',
       'HR',
       NOW()
   )
   RETURNING id;
   
   -- Then insert EmployeeData (use the returned user ID)
   INSERT INTO employee_data (user_id, browser_fingerprint, position, created_at)
   VALUES (
       <user_id_from_above>,
       '396520d70ea1f79dd21caffd85085795',
       'HR Manager',
       NOW()
   );
   
   -- Insert UserStatus
   INSERT INTO user_status (user_id, is_active, failed_attempts)
   VALUES (
       <user_id_from_above>,
       true,
       0
   );
   ```

**Option C: Copy from Local Database**
If you have data in your local SQLite database:

1. Export from local:
   ```bash
   python export_local_data.py
   ```

2. Import to production:
   ```bash
   python import_to_production.py
   ```

---

## Step 6: Verify Everything Works

### 6.1 Check Database Tables
```python
from app import app, db
from sqlalchemy import inspect

with app.app_context():
    inspector = inspect(db.engine)
    tables = inspector.get_table_names()
    print("Tables in database:")
    for table in tables:
        print(f"  - {table}")
```

### 6.2 Check Users
```python
from app import app, db, User, EmployeeData

with app.app_context():
    users = User.query.all()
    print(f"\nTotal users: {len(users)}")
    for user in users:
        emp = EmployeeData.query.filter_by(user_id=user.id).first()
        fp = emp.browser_fingerprint if emp else "None"
        print(f"  - {user.username} ({user.department}) - Fingerprint: {fp[:16]}...")
```

### 6.3 Test Login
1. Visit your app URL: `https://your-app.onrender.com/login`
2. Try logging in with HR Manager credentials
3. Check logs for any errors

---

## Step 7: Common Issues and Solutions

### Issue 1: "No matching fingerprint found. Checked 0 stored fingerprints"
**Solution:**
- Database is empty or fingerprint not stored
- Run Step 5 to add users with fingerprints

### Issue 2: "SSL error: decryption failed or bad record mac"
**Solution:**
- Database connection issue
- Make sure you're using **Internal Database URL** (not External)
- Check that `DATABASE_URL` is set correctly

### Issue 3: "Tables don't exist"
**Solution:**
- Run `db.create_all()` manually (Step 4.2)
- Check logs for migration errors

### Issue 4: "Can't connect to database"
**Solution:**
- Verify `DATABASE_URL` is set in environment variables
- Make sure database is running (check Render dashboard)
- Try using Internal URL instead of External

---

## Step 8: Database Connection URLs Explained

### Internal Database URL
- **Format:** `postgresql://user:password@hostname:5432/database`
- **Use:** For web services on Render (same network)
- **Example:** `postgresql://ngtech_user:abc123@dpg-xxxxx-a.oregon-postgres.render.com/ngtech_db`
- **Access:** Only from Render services

### External Database URL
- **Format:** Similar but with different hostname
- **Use:** For external tools (pgAdmin, DBeaver, local scripts)
- **Access:** From anywhere (requires SSL)

### Which One to Use?
- **For Web Service:** Use **Internal Database URL** (set as `DATABASE_URL`)
- **For Local Tools:** Use **External Database URL** (with SSL)

---

## Step 9: Quick Setup Checklist

- [ ] PostgreSQL database created on Render
- [ ] Database linked to Web Service (or `DATABASE_URL` set manually)
- [ ] Environment variables configured (`SECRET_KEY`, `ENCRYPTION_MASTER_KEY`, etc.)
- [ ] Web service deployed
- [ ] Checked logs: `[OK] Using PostgreSQL database (Production mode)`
- [ ] Tables created (check logs or run `db.create_all()`)
- [ ] Initial users added (HR Manager with fingerprint)
- [ ] Tested login functionality
- [ ] Verified fingerprint matching works

---

## Step 10: Useful Commands

### Check Database Connection
```python
from app import app, db
with app.app_context():
    try:
        db.session.execute(db.text("SELECT 1"))
        print("✅ Database connected!")
    except Exception as e:
        print(f"❌ Error: {e}")
```

### List All Users
```python
from app import app, db, User
with app.app_context():
    for user in User.query.all():
        print(f"{user.username} - {user.department}")
```

### Update Fingerprint
```python
from app import app, db, User, EmployeeData
with app.app_context():
    hr = User.query.filter_by(username='hr_user').first()
    if hr:
        emp = EmployeeData.query.filter_by(user_id=hr.id).first()
        if emp:
            emp.browser_fingerprint = '396520d70ea1f79dd21caffd85085795'
            db.session.commit()
            print("✅ Fingerprint updated!")
```

---

## Need Help?

If you encounter issues:
1. Check Render logs: Web Service → Logs tab
2. Check database status: Database → Metrics tab
3. Verify environment variables are set correctly
4. Test database connection using Step 10 commands

---

**Remember:** Always use the **Internal Database URL** for your web service on Render!

