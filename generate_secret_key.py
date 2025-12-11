"""
Quick script to generate a secure SECRET_KEY for production
Run this and copy the output to your environment variables
"""
import secrets

# Generate a secure random key
secret_key = secrets.token_hex(32)

print("=" * 70)
print("Generated SECRET_KEY for Production")
print("=" * 70)
print()
print(f"SECRET_KEY={secret_key}")
print()
print("Copy this value and set it as an environment variable:")
print("  - Railway: Project → Variables → Add SECRET_KEY")
print("  - Render: Service → Environment → Add SECRET_KEY")
print("  - Heroku: heroku config:set SECRET_KEY='{value}'")
print()
print("=" * 70)
print("⚠️  Keep this key secret! Do not commit it to Git.")
print("=" * 70)

