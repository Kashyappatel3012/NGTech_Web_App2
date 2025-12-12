"""
Test script to verify fingerprint generation and validation
Tests the exact same logic used in generate_fingerprint.html
"""
import sys
import io
import hashlib

# Fix encoding for Windows console
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

def generate_fingerprint_md5(components):
    """Generate MD5 hash from components (same as JavaScript)"""
    fingerprint_string = '|'.join(components)
    # Generate MD5 hash (same as CryptoJS.MD5 in JavaScript)
    md5_hash = hashlib.md5(fingerprint_string.encode('utf-8')).hexdigest()
    return md5_hash

def test_fingerprint_logic():
    """Test the fingerprint generation logic"""
    print("=" * 70)
    print("Fingerprint Generation Test")
    print("=" * 70)
    print("\nTesting fingerprint generation logic...\n")
    
    # Simulate browser components (in the exact order as JavaScript)
    # This should match the order in generate_fingerprint.html
    test_components = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',  # User Agent
        '1920x1080x24',  # Screen Resolution
        'Asia/Kolkata',  # Timezone
        '330',  # Timezone Offset
        'en-US',  # Language
        'en-US,en',  # Languages
        'Win32',  # Platform
        '8',  # Hardware Concurrency
        '8',  # Device Memory
        '0'  # Max Touch Points
    ]
    
    print("Components (in order):")
    for i, comp in enumerate(test_components, 1):
        print(f"  {i}. {comp}")
    
    fingerprint_string = '|'.join(test_components)
    print(f"\nFingerprint String: {fingerprint_string[:100]}...")
    
    # Generate MD5 hash
    fingerprint = generate_fingerprint_md5(test_components)
    print(f"\nGenerated Fingerprint: {fingerprint}")
    print(f"Length: {len(fingerprint)} characters")
    
    # Test with reference fingerprint
    reference_fp = "396520d70ea1f79dd21caffd85085795"
    print(f"\nReference Fingerprint: {reference_fp}")
    print(f"Match: {'YES' if fingerprint == reference_fp else 'NO'}")
    
    if fingerprint != reference_fp:
        print("\n⚠️  Fingerprints don't match!")
        print("This means the browser components are different.")
        print("The fingerprint will vary based on:")
        print("  - User Agent (browser and version)")
        print("  - Screen Resolution")
        print("  - Timezone")
        print("  - Language settings")
        print("  - Platform")
        print("  - Hardware specs")
    
    return fingerprint

def test_database_validation():
    """Test database validation logic"""
    print("\n" + "=" * 70)
    print("Database Validation Test")
    print("=" * 70)
    
    from app import app, db, User, EmployeeData, validate_browser_fingerprint
    
    with app.app_context():
        # Find HR Manager
        hr_user = User.query.filter_by(username='hr_user').first() or User.query.filter_by(department='HR').first()
        
        if not hr_user:
            print("❌ HR Manager user not found!")
            return
        
        print(f"✅ Found HR Manager: {hr_user.username} (ID: {hr_user.id})")
        
        # Get stored fingerprint
        emp_data = EmployeeData.query.filter_by(user_id=hr_user.id).first()
        if not emp_data:
            print("❌ Employee data not found!")
            return
        
        stored_fp = emp_data.browser_fingerprint
        if not stored_fp:
            print("❌ No fingerprint stored in database!")
            return
        
        print(f"\n📋 Stored Fingerprint: {stored_fp}")
        print(f"   Length: {len(stored_fp)} characters")
        
        # Test validation with reference fingerprint
        reference_fp = "396520d70ea1f79dd21caffd85085795"
        print(f"\n🧪 Testing with reference fingerprint: {reference_fp}")
        
        # Test validation with user
        is_valid, user_found = validate_browser_fingerprint(reference_fp, hr_user)
        print(f"   Validation (with user): {is_valid}")
        print(f"   User found: {user_found.username if user_found else 'None'}")
        
        # Test validation without user
        is_valid2, user_found2 = validate_browser_fingerprint(reference_fp, None)
        print(f"   Validation (without user): {is_valid2}")
        print(f"   User found: {user_found2.username if user_found2 else 'None'}")
        
        # Compare stored vs reference
        stored_clean = stored_fp.strip()
        reference_clean = reference_fp.strip()
        
        print(f"\n📊 Comparison:")
        print(f"   Stored (cleaned): {stored_clean}")
        print(f"   Reference (cleaned): {reference_clean}")
        print(f"   Match: {'YES ✅' if stored_clean == reference_clean else 'NO ❌'}")
        
        if stored_clean != reference_clean:
            print(f"\n⚠️  Mismatch detected!")
            print(f"   Stored length: {len(stored_clean)}")
            print(f"   Reference length: {len(reference_clean)}")
            print(f"   First 16 chars - Stored: {stored_clean[:16]}, Reference: {reference_clean[:16]}")
            print(f"\n💡 Solution: Update database with reference fingerprint")
        else:
            print(f"\n✅ Fingerprints match! Validation should work.")

if __name__ == '__main__':
    try:
        # Test fingerprint generation logic
        test_fingerprint_logic()
        
        # Test database validation
        test_database_validation()
        
        print("\n" + "=" * 70)
        print("✅ Test completed!")
        print("=" * 70)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()

