# Production Deployment Guide - Complete PAAS Options

This guide provides **3 different PAAS (Platform as a Service) options** for deploying your Flask application to production with **minimal maintenance required**. All options handle server management, scaling, and infrastructure automatically.

---

## Table of Contents

1. [Pre-Deployment Checklist](#pre-deployment-checklist)
2. [Option 1: Railway (Recommended - Easiest)](#option-1-railway-recommended---easiest)
3. [Option 2: Render (Best Free Tier)](#option-2-render-best-free-tier)
4. [Option 3: Heroku (Most Established)](#option-3-heroku-most-established)
5. [Post-Deployment Configuration](#post-deployment-configuration)
6. [Troubleshooting](#troubleshooting)
7. [Maintenance & Monitoring](#maintenance--monitoring)

---

## Pre-Deployment Checklist

### 1. Code Preparation

#### A. Update Database Configuration for Production

Your `app.py` already supports PostgreSQL via environment variables. However, let's ensure it's production-ready:

**Current configuration (lines 224-237 in app.py):**
```python
# Database configuration with fallback
db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance', 'db.sqlite')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'
```

**Add this BEFORE the SQLite configuration (around line 224):**

```python
# Database configuration with fallback
# Check for production database URL (PAAS providers set this automatically)
database_url = os.environ.get('DATABASE_URL')
if database_url:
    # Convert DATABASE_URL format (postgres://) to SQLAlchemy format (postgresql://)
    if database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql://', 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = database_url
    print("✅ Using PostgreSQL database (Production mode)")
else:
    # Development: Use SQLite
    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'instance', 'db.sqlite')
    if not os.path.exists(os.path.dirname(db_path)):
        os.makedirs(os.path.dirname(db_path), mode=0o700, exist_ok=True)
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{db_path}'
    print("✅ Using SQLite database (Development mode)")
```

#### B. Update Email Configuration for Environment Variables

**Current (lines 254-260):**
```python
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'techumen3012@gmail.com'
app.config['MAIL_PASSWORD'] = 'imso zkvi tdmz rrxu'
app.config['MAIL_DEFAULT_SENDER'] = 'techumen3012@gmail.com'
```

**Replace with:**
```python
# Email configuration - use environment variables in production
app.config['MAIL_SERVER'] = os.environ.get('MAIL_SERVER', 'smtp.gmail.com')
app.config['MAIL_PORT'] = int(os.environ.get('MAIL_PORT', 587))
app.config['MAIL_USE_TLS'] = os.environ.get('MAIL_USE_TLS', 'True').lower() == 'true'
app.config['MAIL_USERNAME'] = os.environ.get('MAIL_USERNAME', 'techumen3012@gmail.com')
app.config['MAIL_PASSWORD'] = os.environ.get('MAIL_PASSWORD', 'imso zkvi tdmz rrxu')
app.config['MAIL_DEFAULT_SENDER'] = os.environ.get('MAIL_DEFAULT_SENDER', 'techumen3012@gmail.com')
```

#### C. Update Secret Key for Production

**Find SECRET_KEY in app.py and replace with:**
```python
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-development-secret-key-change-in-production')
```

**Generate a secure secret key:**
```python
import secrets
print(secrets.token_hex(32))
```

#### D. Create Procfile (for Heroku/Railway)

Create a file named `Procfile` (no extension) in the root directory:

```
web: gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
```

#### E. Create runtime.txt (Optional - specify Python version)

Create `runtime.txt`:
```
python-3.11.0
```

#### F. Update requirements.txt

Add `gunicorn` for production server:

```txt
gunicorn>=21.2.0  # Production WSGI server
```

#### G. Create .gitignore (if not exists)

Create `.gitignore`:
```
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
ENV/
.venv

# Flask
instance/
*.db
*.sqlite
*.sqlite3

# Environment variables
.env
.env.local

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db

# Logs
*.log
app.log

# Uploads (optional - depends on your needs)
# static/uploads/*
# !static/uploads/.gitkeep

# Temporary files
*.tmp
*.temp
remove_inactive_users_one_time.py
```

### 2. Security Checklist

- [ ] Change all hardcoded passwords/keys to environment variables
- [ ] Generate new SECRET_KEY for production
- [ ] Update email credentials (use App Password for Gmail)
- [ ] Ensure FLASK_DEBUG is False in production
- [ ] Review and update CORS settings if needed
- [ ] Enable HTTPS/SSL (automatic on PAAS platforms)

### 3. Test Locally with Production Settings

```bash
# Set environment variables
export DATABASE_URL="postgresql://user:pass@localhost:5432/dbname"
export SECRET_KEY="your-generated-secret-key"
export FLASK_ENV="production"

# Install gunicorn
pip install gunicorn

# Test with gunicorn
gunicorn app:app --bind 0.0.0.0:5000 --workers 2
```

---

## Option 1: Railway (Recommended - Easiest)

**Best for:** Quick deployment, automatic HTTPS, PostgreSQL included, easy scaling

**Pricing:** Free tier available, then $5/month + usage

### Step-by-Step Deployment

#### 1. Sign Up
- Go to https://railway.app
- Sign up with GitHub (recommended) or email

#### 2. Create New Project
- Click "New Project"
- Select "Deploy from GitHub repo" (recommended) or "Empty Project"

#### 3. Connect GitHub Repository
- If using GitHub: Select your repository
- Railway will automatically detect it's a Python/Flask app

#### 4. Add PostgreSQL Database
- In your project dashboard, click "+ New"
- Select "Database" → "PostgreSQL"
- Railway automatically creates database and sets `DATABASE_URL` environment variable

#### 5. Configure Environment Variables

Click on your service → "Variables" tab, add:

```
SECRET_KEY=your-generated-secret-key-here
MAIL_USERNAME=techumen3012@gmail.com
MAIL_PASSWORD=your-gmail-app-password
MAIL_DEFAULT_SENDER=techumen3012@gmail.com
FLASK_ENV=production
PORT=5000
```

**Note:** `DATABASE_URL` is automatically set by Railway when you add PostgreSQL.

#### 6. Configure Build Settings

Railway auto-detects, but you can verify:
- **Build Command:** (leave empty - auto-detected)
- **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`

#### 7. Deploy

- Railway automatically deploys on every push to main branch
- Or click "Deploy" button
- Wait for build to complete (2-5 minutes)

#### 8. Get Your URL

- Railway provides a URL like: `https://your-app-name.up.railway.app`
- You can add custom domain later

#### 9. Run Database Migrations

After first deployment, run migrations:

**Option A: Via Railway CLI**
```bash
# Install Railway CLI
npm i -g @railway/cli

# Login
railway login

# Link to project
railway link

# Run migrations
railway run python app.py
# (This will initialize database tables)
```

**Option B: Via Railway Dashboard**
- Go to your service → "Deployments" → "View Logs"
- Check if database tables are created automatically (they should be via `initialize_database()`)

#### 10. Verify Deployment

- Visit your Railway URL
- Test login functionality
- Check database connection
- Test file uploads

### Railway Advantages

✅ **Automatic HTTPS** - SSL certificates managed automatically  
✅ **Auto-scaling** - Handles traffic spikes  
✅ **PostgreSQL included** - No separate database setup  
✅ **GitHub integration** - Auto-deploy on push  
✅ **Free tier** - Good for testing  
✅ **Easy custom domains** - Add your domain in settings  
✅ **Built-in monitoring** - View logs and metrics  

### Railway Pricing

- **Free:** $5 credit/month (enough for small apps)
- **Hobby:** $5/month + usage
- **Pro:** $20/month + usage

---

## Option 2: Render (Best Free Tier)

**Best for:** Free tier with PostgreSQL, good for small to medium apps

**Pricing:** Free tier available, then $7/month for web service

### Step-by-Step Deployment

#### 1. Sign Up
- Go to https://render.com
- Sign up with GitHub (recommended)

#### 2. Create New Web Service
- Click "New +" → "Web Service"
- Connect your GitHub repository
- Select the repository and branch

#### 3. Configure Service

**Settings:**
- **Name:** your-app-name
- **Environment:** Python 3
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
- **Plan:** Free (or paid for better performance)

#### 4. Add PostgreSQL Database
- Click "New +" → "PostgreSQL"
- **Name:** your-app-db
- **Database:** your_app_db
- **User:** (auto-generated)
- **Region:** Choose closest to you
- **Plan:** Free (or paid)
- **Note the connection string** (Internal Database URL)

#### 5. Link Database to Web Service
- Go to your Web Service → "Environment"
- Add environment variable:
  - **Key:** `DATABASE_URL`
  - **Value:** (Copy from PostgreSQL service → "Connections" → "Internal Database URL")

#### 6. Configure Environment Variables

In Web Service → "Environment", add:

```
SECRET_KEY=your-generated-secret-key-here
MAIL_USERNAME=techumen3012@gmail.com
MAIL_PASSWORD=your-gmail-app-password
MAIL_DEFAULT_SENDER=techumen3012@gmail.com
FLASK_ENV=production
PYTHON_VERSION=3.11.0
```

#### 7. Deploy
- Click "Create Web Service"
- Render will build and deploy (5-10 minutes first time)
- Watch build logs for any errors

#### 8. Get Your URL
- Render provides: `https://your-app-name.onrender.com`
- Free tier: App sleeps after 15 minutes of inactivity (wakes on first request)

#### 9. Run Database Migrations

After deployment, initialize database:

**Via Render Shell:**
- Go to your service → "Shell"
- Run: `python app.py` (this will create tables)

**Or via Logs:**
- Check deployment logs - `initialize_database()` should run automatically

#### 10. Custom Domain (Optional)
- Go to Settings → "Custom Domains"
- Add your domain
- Update DNS records as instructed

### Render Advantages

✅ **Generous free tier** - PostgreSQL + Web service free  
✅ **Auto-deploy from GitHub**  
✅ **Automatic HTTPS**  
✅ **Easy PostgreSQL setup**  
✅ **Good documentation**  
✅ **Sleep mode** - Free tier sleeps but wakes automatically  

### Render Limitations (Free Tier)

⚠️ **Sleep mode** - App sleeps after 15 min inactivity (first request takes ~30 seconds)  
⚠️ **Limited resources** - 512MB RAM, shared CPU  

### Render Pricing

- **Free:** Web service + PostgreSQL (with sleep mode)
- **Starter:** $7/month per service (no sleep, better performance)
- **Standard:** $25/month per service

---

## Option 3: Heroku (Most Established)

**Best for:** Enterprise-grade, most features, established platform

**Pricing:** No free tier (discontinued), starts at $5/month

### Step-by-Step Deployment

#### 1. Sign Up
- Go to https://heroku.com
- Sign up for account

#### 2. Install Heroku CLI
```bash
# Windows: Download from https://devcenter.heroku.com/articles/heroku-cli
# Or use: npm install -g heroku

# Mac
brew tap heroku/brew && brew install heroku

# Linux
curl https://cli-assets.heroku.com/install.sh | sh
```

#### 3. Login
```bash
heroku login
```

#### 4. Create Heroku App
```bash
# Navigate to your project directory
cd /path/to/your/project

# Create app
heroku create your-app-name

# Or create via dashboard: https://dashboard.heroku.com/new-app
```

#### 5. Add PostgreSQL Database
```bash
# Add PostgreSQL addon (free tier available)
heroku addons:create heroku-postgresql:mini

# This automatically sets DATABASE_URL environment variable
```

#### 6. Set Environment Variables
```bash
heroku config:set SECRET_KEY="your-generated-secret-key-here"
heroku config:set MAIL_USERNAME="techumen3012@gmail.com"
heroku config:set MAIL_PASSWORD="your-gmail-app-password"
heroku config:set MAIL_DEFAULT_SENDER="techumen3012@gmail.com"
heroku config:set FLASK_ENV="production"
```

#### 7. Deploy
```bash
# If using Git (recommended)
git init  # if not already a git repo
git add .
git commit -m "Initial commit"
git push heroku main

# Or connect GitHub repository in Heroku dashboard
```

#### 8. Run Database Migrations
```bash
# After first deployment
heroku run python app.py
# Or
heroku run flask db upgrade  # if using Flask-Migrate
```

#### 9. Open Your App
```bash
heroku open
# Or visit: https://your-app-name.herokuapp.com
```

#### 10. View Logs
```bash
heroku logs --tail
```

### Heroku Advantages

✅ **Most established platform** - 15+ years  
✅ **Excellent documentation**  
✅ **Many addons available**  
✅ **Automatic HTTPS**  
✅ **Easy scaling** - `heroku ps:scale web=2`  
✅ **GitHub integration**  
✅ **CLI tools** - Powerful command-line interface  

### Heroku Pricing

- **Eco:** $5/month (sleeps after 30 min)
- **Basic:** $7/month (always on)
- **Standard:** $25/month (better performance)

---

## Post-Deployment Configuration

### 1. Gmail App Password Setup

For email functionality, create a Gmail App Password:

1. Go to https://myaccount.google.com/
2. Security → 2-Step Verification (enable if not enabled)
3. App passwords → Generate new app password
4. Use this password in `MAIL_PASSWORD` environment variable

### 2. Custom Domain Setup

**Railway:**
- Settings → "Domains" → Add custom domain
- Update DNS records as shown

**Render:**
- Settings → "Custom Domains" → Add domain
- Follow DNS configuration instructions

**Heroku:**
```bash
heroku domains:add www.yourdomain.com
heroku domains:add yourdomain.com
# Then update DNS records
```

### 3. Database Backup Setup

**Railway:**
- PostgreSQL service → "Backups" → Enable automatic backups

**Render:**
- PostgreSQL service → "Backups" → Schedule backups

**Heroku:**
```bash
heroku addons:create pgbackups:auto-month
```

### 4. Monitoring & Logs

All platforms provide:
- **Logs:** View real-time application logs
- **Metrics:** CPU, memory, request metrics
- **Alerts:** Set up email alerts for errors

### 5. Scheduled Tasks (Email Scheduler)

Your `daily_workplan_email_scheduler.py` uses APScheduler. On PAAS platforms:

**Option A: Keep as-is** (works on all platforms)
- Scheduler runs in the same process
- Works fine for single-instance deployments

**Option B: Use Platform Schedulers** (for multi-instance)

**Railway:**
- Use Railway Cron Jobs (addon)

**Render:**
- Use Render Cron Jobs (addon)

**Heroku:**
- Use Heroku Scheduler addon:
```bash
heroku addons:create scheduler:standard
# Then configure in dashboard: https://dashboard.heroku.com/scheduler
```

### 6. Static Files & Uploads

**Important:** PAAS platforms have **ephemeral filesystems** - files are deleted on restart.

**Solution:** Use cloud storage for uploads:

**Option A: AWS S3** (Recommended)
```python
# Install boto3
pip install boto3

# Update upload handling to use S3
```

**Option B: Cloudinary** (Easy)
```bash
pip install cloudinary
```

**Option C: Keep on filesystem** (Temporary)
- Files will be lost on restart
- Only for development/testing

---

## Troubleshooting

### Issue: Database Connection Failed

**Symptoms:** `could not connect to server`

**Solutions:**
1. Check `DATABASE_URL` environment variable is set
2. Verify database service is running
3. Check connection string format
4. Ensure database is in same region as app

### Issue: Application Crashes on Startup

**Symptoms:** App fails to start, shows error in logs

**Solutions:**
1. Check logs: `railway logs` / `render logs` / `heroku logs`
2. Verify all environment variables are set
3. Check `requirements.txt` has all dependencies
4. Ensure `Procfile` start command is correct
5. Verify Python version compatibility

### Issue: Static Files Not Loading

**Symptoms:** CSS/JS/images not loading

**Solutions:**
1. Check `static` folder is in repository
2. Verify file paths are relative (not absolute)
3. Check Flask static folder configuration
4. Clear browser cache

### Issue: Email Not Sending

**Symptoms:** Emails not being sent

**Solutions:**
1. Verify Gmail App Password (not regular password)
2. Check `MAIL_USERNAME` and `MAIL_PASSWORD` are set
3. Check firewall/network restrictions
4. Review email logs in application logs

### Issue: Database Tables Not Created

**Symptoms:** 500 errors, table not found

**Solutions:**
1. Run migrations manually:
   ```bash
   # Railway
   railway run python app.py
   
   # Render
   render shell → python app.py
   
   # Heroku
   heroku run python app.py
   ```
2. Check `initialize_database()` function runs on startup
3. Verify database permissions

### Issue: File Uploads Not Working

**Symptoms:** Uploads fail or files disappear

**Solutions:**
1. **Use cloud storage** (S3, Cloudinary) - filesystem is ephemeral
2. Check file size limits
3. Verify upload directory permissions
4. Check `MAX_CONTENT_LENGTH` setting

---

## Maintenance & Monitoring

### Regular Tasks

1. **Monitor Logs** - Check for errors weekly
2. **Database Backups** - Verify backups are running
3. **Update Dependencies** - Keep packages updated
4. **Security Updates** - Monitor for vulnerabilities
5. **Performance Monitoring** - Check metrics regularly

### Scaling

**When to Scale:**
- High CPU usage (>80%)
- High memory usage (>80%)
- Slow response times
- Increased traffic

**How to Scale:**

**Railway:**
- Settings → "Scaling" → Increase instances

**Render:**
- Settings → "Plan" → Upgrade plan

**Heroku:**
```bash
heroku ps:scale web=2  # Scale to 2 dynos
```

### Cost Optimization

1. **Use free tiers** for development/testing
2. **Monitor usage** - Set up billing alerts
3. **Optimize database queries** - Reduce database calls
4. **Use CDN** for static files (reduces server load)
5. **Enable caching** where possible

---

## Comparison Table

| Feature | Railway | Render | Heroku |
|---------|---------|--------|--------|
| **Free Tier** | $5 credit/month | Yes (with sleep) | No |
| **PostgreSQL** | Included | Included | Addon |
| **Auto HTTPS** | ✅ | ✅ | ✅ |
| **GitHub Deploy** | ✅ | ✅ | ✅ |
| **Ease of Use** | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| **Documentation** | Good | Excellent | Excellent |
| **Scaling** | Easy | Easy | Easy |
| **Custom Domain** | ✅ | ✅ | ✅ |
| **CLI Tools** | Good | Limited | Excellent |
| **Best For** | Quick start | Free tier | Enterprise |

---

## Recommended Choice

**For Most Users:** **Railway** - Easiest setup, good free tier, PostgreSQL included

**For Free Tier:** **Render** - Best free tier with PostgreSQL

**For Enterprise:** **Heroku** - Most features, best documentation

---

## Quick Start Commands

### Railway
```bash
# Install CLI
npm i -g @railway/cli

# Login
railway login

# Deploy
railway up
```

### Render
```bash
# Just connect GitHub repo in dashboard
# No CLI needed for basic deployment
```

### Heroku
```bash
# Install CLI
# Login
heroku login

# Deploy
git push heroku main
```

---

## Security Checklist

- [ ] SECRET_KEY is strong and unique
- [ ] All passwords in environment variables (not code)
- [ ] FLASK_DEBUG=False in production
- [ ] HTTPS enabled (automatic on all platforms)
- [ ] Database credentials secure
- [ ] Email credentials use App Password
- [ ] CORS configured properly
- [ ] File upload limits set
- [ ] Session security enabled
- [ ] Regular security updates

---

## Support & Resources

- **Railway Docs:** https://docs.railway.app
- **Render Docs:** https://render.com/docs
- **Heroku Docs:** https://devcenter.heroku.com
- **Flask Production:** https://flask.palletsprojects.com/en/latest/deploying/

---

**Last Updated:** December 2025  
**Application:** NTP33 Flask Web Application

