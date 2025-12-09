# Production Database Setup Guide

## Quick Start: Switch to PostgreSQL Production Database

Your code is already configured to automatically use PostgreSQL when environment variables are set. **No code changes needed!**

## Step-by-Step Instructions

### Step 1: Install PostgreSQL (if not already installed)

**Windows:**
- Download from: https://www.postgresql.org/download/windows/
- Install using the installer
- Remember the password you set for the `postgres` user

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install postgresql postgresql-contrib
```

**macOS:**
```bash
brew install postgresql
brew services start postgresql
```

### Step 2: Create the Database

```bash
# Connect to PostgreSQL
psql -U postgres

# Create the database
CREATE DATABASE ntp3_db;

# Create a dedicated user (optional but recommended)
CREATE USER ntp3_user WITH PASSWORD 'your_secure_password';
GRANT ALL PRIVILEGES ON DATABASE ntp3_db TO ntp3_user;

# Exit psql
\q
```

### Step 3: Install PostgreSQL Driver

```bash
pip install psycopg2-binary
```

Or install all requirements:
```bash
pip install -r requirements.txt
```

### Step 4: Set Environment Variables

Choose **ONE** of the following methods:

#### Method 1: Using DATABASE_URL (Recommended - Single Variable)

**Windows (PowerShell):**
```powershell
$env:DATABASE_URL="postgresql://ntp3_user:your_secure_password@localhost:5432/ntp3_db"
```

**Windows (Command Prompt):**
```cmd
set DATABASE_URL=postgresql://ntp3_user:your_secure_password@localhost:5432/ntp3_db
```

**Linux/macOS:**
```bash
export DATABASE_URL="postgresql://ntp3_user:your_secure_password@localhost:5432/ntp3_db"
```

#### Method 2: Using Individual Variables

**Windows (PowerShell):**
```powershell
$env:FLASK_ENV="production"
$env:DB_HOST="localhost"
$env:DB_PORT="5432"
$env:DB_NAME="ntp3_db"
$env:DB_USER="ntp3_user"
$env:DB_PASSWORD="your_secure_password"
```

**Windows (Command Prompt):**
```cmd
set FLASK_ENV=production
set DB_HOST=localhost
set DB_PORT=5432
set DB_NAME=ntp3_db
set DB_USER=ntp3_user
set DB_PASSWORD=your_secure_password
```

**Linux/macOS:**
```bash
export FLASK_ENV=production
export DB_HOST=localhost
export DB_PORT=5432
export DB_NAME=ntp3_db
export DB_USER=ntp3_user
export DB_PASSWORD=your_secure_password
```

### Step 5: Run Your Application

```bash
python app.py
```

You should see:
```
✅ Using PostgreSQL database (Production mode)
✅ PostgreSQL connection successful
✅ Database tables created/verified successfully
```

## For Production Servers (Permanent Setup)

### Option A: Using .env file (Recommended)

1. Create a `.env` file in your project root:
```env
DATABASE_URL=postgresql://ntp3_user:your_secure_password@localhost:5432/ntp3_db
FLASK_ENV=production
SECRET_KEY=your-production-secret-key-here
```

2. Install python-dotenv:
```bash
pip install python-dotenv
```

3. Add to the top of `app.py` (after imports):
```python
from dotenv import load_dotenv
load_dotenv()  # Load environment variables from .env file
```

### Option B: System Environment Variables

**Linux (systemd service):**
Add to `/etc/systemd/system/your-app.service`:
```ini
[Service]
Environment="DATABASE_URL=postgresql://user:pass@host:5432/dbname"
Environment="FLASK_ENV=production"
```

**Windows (Service):**
Set environment variables in the service configuration.

**Docker:**
Add to `docker-compose.yml`:
```yaml
environment:
  - DATABASE_URL=postgresql://user:pass@db:5432/dbname
  - FLASK_ENV=production
```

## Verification

After setting up, verify the connection:

1. **Check console output:**
   - Should show "✅ Using PostgreSQL database (Production mode)"
   - Should show "✅ PostgreSQL connection successful"

2. **Check database:**
```bash
psql -U ntp3_user -d ntp3_db -c "\dt"
```
This should list all your tables.

## Troubleshooting

### Error: "could not connect to server"
- **Solution:** Make sure PostgreSQL is running:
  - Windows: Check Services, start "postgresql-x64-XX" service
  - Linux: `sudo systemctl start postgresql`
  - macOS: `brew services start postgresql`

### Error: "password authentication failed"
- **Solution:** Check your password in environment variables matches PostgreSQL user password

### Error: "database does not exist"
- **Solution:** Create the database: `createdb ntp3_db` or use SQL: `CREATE DATABASE ntp3_db;`

### Error: "permission denied"
- **Solution:** Grant privileges:
```sql
GRANT ALL PRIVILEGES ON DATABASE ntp3_db TO ntp3_user;
```

### Still using SQLite?
- Check environment variables are set correctly
- Restart your application after setting environment variables
- On Windows, make sure you're using the same terminal session

## Switching Back to Development (SQLite)

Simply **unset** or **remove** the environment variables:

**Windows (PowerShell):**
```powershell
Remove-Item Env:\DATABASE_URL
Remove-Item Env:\FLASK_ENV
```

**Windows (Command Prompt):**
```cmd
set DATABASE_URL=
set FLASK_ENV=
```

**Linux/macOS:**
```bash
unset DATABASE_URL
unset FLASK_ENV
```

## Production Checklist

- [ ] PostgreSQL installed and running
- [ ] Database created (`ntp3_db`)
- [ ] Database user created with proper permissions
- [ ] `psycopg2-binary` installed
- [ ] Environment variables set (DATABASE_URL or individual variables)
- [ ] Application starts and shows "✅ Using PostgreSQL database"
- [ ] Tables created successfully
- [ ] Can connect to database and query tables
- [ ] Backup strategy in place
- [ ] Connection pooling configured (already done in code)

## Security Notes

1. **Never commit `.env` file** - Add to `.gitignore`
2. **Use strong passwords** for database users
3. **Restrict database access** - Only allow connections from your application server
4. **Use SSL connections** in production (add `?sslmode=require` to DATABASE_URL)
5. **Rotate secrets regularly** - Change passwords and SECRET_KEY periodically

## Example Production DATABASE_URL

```bash
# With SSL (recommended for production)
DATABASE_URL="postgresql://user:pass@host:5432/dbname?sslmode=require"

# Remote server
DATABASE_URL="postgresql://user:pass@your-server.com:5432/dbname"

# With connection pool settings
DATABASE_URL="postgresql://user:pass@host:5432/dbname?pool_size=10&max_overflow=20"
```

---

**Remember:** Your code automatically detects the environment and switches between SQLite (development) and PostgreSQL (production). Just set the environment variables and you're ready to go!

