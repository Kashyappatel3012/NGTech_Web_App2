# Render.com Environment Variables Guide

## Required Environment Variables for Production

### 🔐 Critical (Must Set)

#### 1. **ENCRYPTION_MASTER_KEY** ⚠️ IMPORTANT
- **Purpose:** Master key for encrypting sensitive data (browser fingerprints in sessions, etc.)
- **Required:** Yes (for production)
- **How to generate:**
  ```bash
  python generate_encryption_key.py
  ```
  Or manually:
  ```python
  import secrets
  print(secrets.token_urlsafe(32))
  ```
- **Example:** `3JLqS0C5KFPy-4-fjxrdNsc-GhkwqfdD7fWTnJcsLrQ`
- **Length:** 43 characters (URL-safe base64)
- **Security:** Keep this SECRET! Never commit to Git.
- **What happens if not set:** App will use auto-generated key (changes on restart, causes data loss)

#### 2. **SECRET_KEY** ⚠️ IMPORTANT
- **Purpose:** Flask session encryption and CSRF protection
- **Required:** Yes (for production)
- **How to generate:**
  ```bash
  python generate_secret_key.py
  ```
  Or manually:
  ```python
  import secrets
  print(secrets.token_hex(32))
  ```
- **Example:** `a1b2c3d4e5f6...` (64 character hex string)
- **Security:** Keep this SECRET! Never commit to Git.
- **What happens if not set:** Uses default key (insecure for production)

#### 3. **DATABASE_URL** ✅ Auto-Set
- **Purpose:** PostgreSQL database connection string
- **Required:** Yes (automatically set by Render when you add PostgreSQL)
- **Format:** `postgresql://user:password@host:port/database`
- **Action:** Render sets this automatically - you don't need to set it manually

### 📧 Email Configuration (Optional but Recommended)

#### 4. **MAIL_USERNAME**
- **Purpose:** Email address for sending emails
- **Required:** No (has default: `techumen3012@gmail.com`)
- **Example:** `techumen3012@gmail.com`

#### 5. **MAIL_PASSWORD**
- **Purpose:** Gmail App Password (not regular password)
- **Required:** No (has default, but should set for production)
- **How to get:** 
  1. Go to Google Account → Security
  2. Enable 2-Step Verification
  3. Generate App Password
  4. Use the 16-character password
- **Example:** `imso zkvi tdmz rrxu`

#### 6. **MAIL_SERVER**
- **Purpose:** SMTP server address
- **Required:** No (default: `smtp.gmail.com`)
- **Example:** `smtp.gmail.com`

#### 7. **MAIL_PORT**
- **Purpose:** SMTP port
- **Required:** No (default: `587`)
- **Example:** `587`

#### 8. **MAIL_USE_TLS**
- **Purpose:** Enable TLS for email
- **Required:** No (default: `True`)
- **Example:** `True`

#### 9. **MAIL_DEFAULT_SENDER**
- **Purpose:** Default sender email address
- **Required:** No (default: `techumen3012@gmail.com`)
- **Example:** `techumen3012@gmail.com`

### ⚙️ Optional Configuration

#### 10. **FLASK_DEBUG**
- **Purpose:** Enable/disable Flask debug mode
- **Required:** No (default: `False`)
- **Value:** `False` (for production)
- **Warning:** Never set to `True` in production!

#### 11. **SESSION_COOKIE_SECURE**
- **Purpose:** Only send cookies over HTTPS
- **Required:** No (default: `False`, but should be `True` for production)
- **Value:** `True` (for HTTPS sites)

#### 12. **USE_HTTPS**
- **Purpose:** Enable HTTPS-specific features
- **Required:** No (default: `False`)
- **Value:** `True` (for production)

#### 13. **PYTHON_VERSION** (Render-specific)
- **Purpose:** Python version to use
- **Required:** No (Render auto-detects)
- **Example:** `3.11.0`

## 📋 Complete Environment Variables List for Render.com

Go to **Render.com Dashboard → Your Web Service → Environment** and add:

```bash
# Critical - MUST SET
ENCRYPTION_MASTER_KEY=your-43-character-encryption-key-here
SECRET_KEY=your-64-character-secret-key-here

# Database - Auto-set by Render (don't set manually)
# DATABASE_URL=postgresql://... (automatically set)

# Email Configuration (recommended to set)
MAIL_USERNAME=techumen3012@gmail.com
MAIL_PASSWORD=your-gmail-app-password-here
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USE_TLS=True
MAIL_DEFAULT_SENDER=techumen3012@gmail.com

# Production Settings
FLASK_DEBUG=False
SESSION_COOKIE_SECURE=True
USE_HTTPS=True
PYTHON_VERSION=3.11.0
```

## 🔧 How to Set Environment Variables on Render.com

### Step-by-Step:

1. **Go to Render.com Dashboard**
   - https://dashboard.render.com

2. **Select Your Web Service**
   - Click on your web service name

3. **Go to "Environment" Tab**
   - Click on "Environment" in the left sidebar

4. **Add Environment Variables**
   - Click "Add Environment Variable"
   - Enter Key and Value
   - Click "Save Changes"

5. **Redeploy**
   - Render will automatically redeploy when you save environment variables
   - Or manually trigger redeploy from "Manual Deploy" tab

## 🎯 Minimum Required for Production

**At minimum, you MUST set these 2:**

1. ✅ `ENCRYPTION_MASTER_KEY` - Prevents encryption key warnings and data loss
2. ✅ `SECRET_KEY` - Required for secure sessions

**Highly Recommended:**

3. ✅ `MAIL_PASSWORD` - For email functionality
4. ✅ `SESSION_COOKIE_SECURE=True` - For HTTPS security
5. ✅ `USE_HTTPS=True` - For HTTPS features

## ⚠️ Security Best Practices

1. **Never commit secrets to Git**
   - All these values should be in environment variables only
   - Add `.env` to `.gitignore`

2. **Use different keys for different environments**
   - Development, Staging, Production should have different keys

3. **Rotate keys periodically**
   - Change keys every 6-12 months for security

4. **Backup your keys securely**
   - Store in password manager
   - If you lose `ENCRYPTION_MASTER_KEY`, encrypted data cannot be recovered

## 🔍 How to Verify Environment Variables

After setting variables, check Render.com logs:

1. Go to your service → "Logs"
2. Look for:
   - ✅ `Using PostgreSQL database (Production mode)` - DATABASE_URL is set
   - ✅ No `WARNING: Using auto-generated encryption key` - ENCRYPTION_MASTER_KEY is set
   - ✅ No security warnings about SECRET_KEY

## 📝 Quick Setup Script

Run these commands locally to generate keys:

```bash
# Generate encryption key
python generate_encryption_key.py

# Generate secret key
python generate_secret_key.py
```

Then copy the generated keys to Render.com environment variables.

## 🆘 Troubleshooting

### Issue: "WARNING: Using auto-generated encryption key"
**Solution:** Set `ENCRYPTION_MASTER_KEY` environment variable

### Issue: Session data lost on restart
**Solution:** Set `SECRET_KEY` environment variable (and `ENCRYPTION_MASTER_KEY`)

### Issue: Email not sending
**Solution:** Check `MAIL_USERNAME` and `MAIL_PASSWORD` are set correctly

### Issue: Database connection failed
**Solution:** Verify `DATABASE_URL` is set (should be automatic from Render PostgreSQL)

