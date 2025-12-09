# Encryption Setup Guide

## Overview
This application uses **AES-256-GCM** (Advanced Encryption Standard with Galois/Counter Mode) for strong encryption of sensitive data. This provides:
- **256-bit encryption** (military-grade security)
- **Authenticated encryption** (prevents tampering)
- **Secure key derivation** using PBKDF2 with 100,000 iterations

## Security Features

### Encryption Algorithm: AES-256-GCM
- **Key Size**: 256 bits (32 bytes)
- **Mode**: GCM (Galois/Counter Mode) - provides both encryption and authentication
- **Nonce**: 12 bytes (96 bits) - randomly generated for each encryption
- **Key Derivation**: PBKDF2-HMAC-SHA256 with 100,000 iterations

### What is Encrypted

1. **Browser Fingerprints** (Database)
   - Stored encrypted in `employee_data.browser_fingerprint`
   - Decrypted only for validation and API responses

2. **Session Data**
   - Browser fingerprints in Flask session
   - Validated fingerprints in session

3. **Data in Transit**
   - Sensitive form data encrypted before storage
   - API responses with sensitive data

## Setup Instructions

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

This will install `cryptography>=41.0.0` which provides AES-256-GCM encryption.

### 2. Set Encryption Key (IMPORTANT!)

**For Production:**
Set the `ENCRYPTION_MASTER_KEY` environment variable with a strong, random key:

**Windows (PowerShell):**
```powershell
$env:ENCRYPTION_MASTER_KEY="your-very-long-random-secure-key-here-minimum-32-characters"
```

**Windows (Command Prompt):**
```cmd
set ENCRYPTION_MASTER_KEY=your-very-long-random-secure-key-here-minimum-32-characters
```

**Linux/macOS:**
```bash
export ENCRYPTION_MASTER_KEY="your-very-long-random-secure-key-here-minimum-32-characters"
```

**Generate a Secure Key:**
```python
import secrets
print(secrets.token_urlsafe(32))  # Generates a 43-character URL-safe key
```

### 3. Encrypt Existing Data

If you have existing browser fingerprints in the database, run the migration script:

```bash
python encrypt_existing_data.py
```

This will encrypt all existing browser fingerprints in the database.

### 4. Verify Encryption

After setup, verify encryption is working:
1. Create a new user with browser fingerprint
2. Check database - fingerprint should be encrypted (long base64 string)
3. Login should work correctly (decryption happens automatically)

## Security Best Practices

1. **Never commit the encryption key to version control**
   - Add `ENCRYPTION_MASTER_KEY` to `.gitignore`
   - Use environment variables or secure key management

2. **Use different keys for different environments**
   - Development, Staging, Production should have different keys

3. **Backup the encryption key securely**
   - If you lose the key, encrypted data cannot be recovered
   - Store backup in secure location (password manager, secure vault)

4. **Rotate keys periodically** (if needed)
   - Create new key
   - Re-encrypt all data with new key
   - Update environment variable

5. **Use HTTPS in production**
   - Encryption protects data at rest
   - HTTPS protects data in transit

## How It Works

### Encryption Flow:
1. **Data Input** → Plaintext browser fingerprint
2. **Key Derivation** → Master key → PBKDF2 → Encryption key (256 bits)
3. **Encryption** → AES-256-GCM encrypts with random nonce
4. **Storage** → Base64-encoded encrypted data stored in database

### Decryption Flow:
1. **Retrieve** → Encrypted data from database
2. **Decode** → Base64 decode
3. **Extract** → Nonce (first 12 bytes) and ciphertext
4. **Decrypt** → AES-256-GCM decrypts and verifies authenticity
5. **Return** → Plaintext browser fingerprint

## Troubleshooting

### Error: "Failed to decrypt data"
- **Cause**: Encryption key mismatch or corrupted data
- **Solution**: Verify `ENCRYPTION_MASTER_KEY` is set correctly

### Warning: "Using auto-generated encryption key"
- **Cause**: `ENCRYPTION_MASTER_KEY` not set
- **Solution**: Set environment variable (see Setup Instructions)

### Data appears encrypted but login fails
- **Cause**: Key changed or data corruption
- **Solution**: 
  1. Verify encryption key is correct
  2. Re-encrypt data if needed
  3. Check database for corruption

## Technical Details

### Key Derivation
- **Algorithm**: PBKDF2-HMAC-SHA256
- **Iterations**: 100,000 (high security)
- **Salt**: Application-specific (derived from app name)
- **Output**: 32 bytes (256 bits)

### Encryption Parameters
- **Algorithm**: AES-256-GCM
- **Key Size**: 256 bits
- **Nonce Size**: 96 bits (12 bytes)
- **Tag Size**: 128 bits (authentication tag)
- **Output Format**: Base64 URL-safe encoding

### Performance
- **Encryption**: ~1-2ms per operation
- **Decryption**: ~1-2ms per operation
- **Impact**: Negligible on application performance

## Support

For encryption-related issues:
1. Check encryption key is set correctly
2. Verify `cryptography` package is installed
3. Check application logs for decryption errors
4. Ensure database column size is sufficient (500 chars for encrypted data)

