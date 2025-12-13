# Render.com Environment Variables - Quick Setup

## ✅ Required Environment Variables

You **MUST** set these 2 environment variables on Render.com:

### 1. ENCRYPTION_MASTER_KEY ⚠️ CRITICAL

**Generated Key:**
```
l3Xr34DMdFx497gno1TcQAynuGREdTuTl65Lz920NdA
```

**Why it's needed:**
- Encrypts sensitive session data (browser fingerprints)
- Without it, you'll see: `WARNING: Using auto-generated encryption key`
- Auto-generated keys change on restart → data loss

### 2. SECRET_KEY ⚠️ CRITICAL

**Generated Key:**
```
321fee1a111d333674c7e8b7fb95d82831151a6e45e57a4000c048c16930c9b7
```

**Why it's needed:**
- Flask session encryption
- CSRF token generation
- Without it, sessions won't work properly

## 📋 Complete Environment Variables List

Go to **Render.com → Your Web Service → Environment** and add:

```bash
# ============================================
# CRITICAL - MUST SET THESE 2
# ============================================
ENCRYPTION_MASTER_KEY=l3Xr34DMdFx497gno1TcQAynuGREdTuTl65Lz920NdA
SECRET_KEY=321fee1a111d333674c7e8b7fb95d82831151a6e45e57a4000c048c16930c9b7

# ============================================
# DATABASE - Auto-set by Render (don't set manually)
# ============================================
# DATABASE_URL is automatically set when you add PostgreSQL

# ============================================
# EMAIL CONFIGURATION (Optional but Recommended)
# ============================================
MAIL_USERNAME=techumen3012@gmail.com
MAIL_PASSWORD=imso zkvi tdmz rrxu
MAIL_SERVER=smtp.gmail.com
MAIL_PORT=587
MAIL_USE_TLS=True
MAIL_DEFAULT_SENDER=techumen3012@gmail.com

# ============================================
# PRODUCTION SETTINGS (Recommended)
# ============================================
FLASK_DEBUG=False
SESSION_COOKIE_SECURE=True
USE_HTTPS=True
PYTHON_VERSION=3.11.0
```

## 🚀 How to Set on Render.com

1. **Go to:** https://dashboard.render.com
2. **Select:** Your web service
3. **Click:** "Environment" tab
4. **Add each variable:**
   - Click "Add Environment Variable"
   - Enter Key: `ENCRYPTION_MASTER_KEY`
   - Enter Value: `l3Xr34DMdFx497gno1TcQAynuGREdTuTl65Lz920NdA`
   - Click "Save"
   - Repeat for `SECRET_KEY` and others
5. **Redeploy:** Render will auto-redeploy after saving

## ⚠️ Important Notes

1. **Keep keys SECRET** - Never commit to Git
2. **Use same keys** - Don't change them after setting (or you'll lose encrypted data)
3. **Backup keys** - Store in password manager
4. **DATABASE_URL** - Automatically set by Render (don't set manually)

## ✅ Verification

After setting, check Render.com logs. You should see:
- ✅ No `WARNING: Using auto-generated encryption key`
- ✅ `Using PostgreSQL database (Production mode)`
- ✅ App starts without errors

## 🔄 If You Need New Keys

**Generate new encryption key:**
```bash
python -c "import secrets; print(secrets.token_urlsafe(32))"
```

**Generate new secret key:**
```bash
python generate_secret_key.py
```

**Note:** Changing keys will cause:
- Loss of encrypted session data
- Users may need to re-login
- Browser fingerprints in sessions will be lost (but database fingerprints remain)

