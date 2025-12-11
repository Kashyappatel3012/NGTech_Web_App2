# Quick Deployment Checklist

## Pre-Deployment (5 minutes)

- [ ] Generate SECRET_KEY: `python -c "import secrets; print(secrets.token_hex(32))"`
- [ ] Create Gmail App Password (for email functionality)
- [ ] Test locally with production settings
- [ ] Commit all changes to Git

## Choose Your Platform

### Option 1: Railway (Recommended)
- [ ] Sign up at https://railway.app
- [ ] Create new project
- [ ] Connect GitHub repository
- [ ] Add PostgreSQL database
- [ ] Set environment variables (see below)
- [ ] Deploy

### Option 2: Render
- [ ] Sign up at https://render.com
- [ ] Create Web Service
- [ ] Create PostgreSQL database
- [ ] Set environment variables
- [ ] Deploy

### Option 3: Heroku
- [ ] Sign up at https://heroku.com
- [ ] Install Heroku CLI
- [ ] Create app: `heroku create your-app-name`
- [ ] Add PostgreSQL: `heroku addons:create heroku-postgresql:mini`
- [ ] Set environment variables
- [ ] Deploy: `git push heroku main`

## Required Environment Variables

Set these in your platform's dashboard:

```
SECRET_KEY=your-generated-secret-key-here
MAIL_USERNAME=techumen3012@gmail.com
MAIL_PASSWORD=your-gmail-app-password
MAIL_DEFAULT_SENDER=techumen3012@gmail.com
FLASK_ENV=production
```

**Note:** `DATABASE_URL` is automatically set when you add PostgreSQL.

## Post-Deployment

- [ ] Verify app is running
- [ ] Test login functionality
- [ ] Check database connection
- [ ] Test file uploads (if using cloud storage)
- [ ] Verify email sending
- [ ] Set up custom domain (optional)
- [ ] Configure backups

## Files Created for Deployment

✅ `Procfile` - Tells platform how to run your app  
✅ `runtime.txt` - Specifies Python version  
✅ `.gitignore` - Excludes sensitive files  
✅ `requirements.txt` - Updated with gunicorn  

## Code Updates Made

✅ Database configuration - Auto-detects PostgreSQL from DATABASE_URL  
✅ Email configuration - Uses environment variables  
✅ SECRET_KEY - Uses environment variable  

## Need Help?

See `PRODUCTION_DEPLOYMENT_GUIDE.md` for detailed instructions.

