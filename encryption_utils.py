"""
Advanced Encryption/Decryption Utility Module
Uses AES-256-GCM (Galois/Counter Mode) for authenticated encryption
Provides strong security with integrity verification
"""
import os
import base64
import hashlib
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.backends import default_backend
import secrets

class EncryptionManager:
    """
    Advanced encryption manager using AES-256-GCM
    Provides authenticated encryption with associated data (AEAD)
    """
    
    def __init__(self, master_key=None):
        """
        Initialize encryption manager
        
        Args:
            master_key: Optional master key. If not provided, will use environment variable
                       or generate a new one (for development only)
        """
        if master_key:
            self.master_key = master_key
        else:
            # Try to get from environment variable first
            self.master_key = os.environ.get('ENCRYPTION_MASTER_KEY')
            if not self.master_key:
                # For production, this should be set via environment variable
                # For development, generate a key (WARNING: This will change on restart)
                try:
                    print("WARNING: Using auto-generated encryption key. Set ENCRYPTION_MASTER_KEY environment variable for production!")
                except:
                    pass  # Ignore encoding errors in print
                self.master_key = self._generate_master_key()
        
        # Derive encryption key from master key using PBKDF2
        self.encryption_key = self._derive_key(self.master_key)
    
    def _generate_master_key(self):
        """Generate a secure random master key (32 bytes = 256 bits)"""
        return secrets.token_urlsafe(32)
    
    def _derive_key(self, master_key, salt=None):
        """
        Derive encryption key from master key using PBKDF2
        
        Args:
            master_key: Master key string
            salt: Optional salt (will generate if not provided)
        
        Returns:
            bytes: Derived encryption key (32 bytes for AES-256)
        """
        if salt is None:
            # Use a fixed salt derived from application name for consistency
            # In production, consider using a stored salt per application instance
            salt = hashlib.sha256(b'NGTech_Assurance_Encryption_Salt_v1').digest()[:16]
        
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,  # 32 bytes = 256 bits for AES-256
            salt=salt,
            iterations=100000,  # High iteration count for security
            backend=default_backend()
        )
        
        # Convert master key to bytes if it's a string
        if isinstance(master_key, str):
            master_key_bytes = master_key.encode('utf-8')
        else:
            master_key_bytes = master_key
        
        key = kdf.derive(master_key_bytes)
        return key
    
    def encrypt(self, plaintext, associated_data=None):
        """
        Encrypt plaintext using AES-256-GCM
        
        Args:
            plaintext: String or bytes to encrypt
            associated_data: Optional associated data for authentication (not encrypted, but authenticated)
        
        Returns:
            str: Base64-encoded encrypted data with nonce (format: nonce:encrypted_data)
        """
        if plaintext is None:
            return None
        
        # Convert to bytes if string
        if isinstance(plaintext, str):
            plaintext_bytes = plaintext.encode('utf-8')
        else:
            plaintext_bytes = plaintext
        
        # Generate random nonce (12 bytes for GCM)
        nonce = secrets.token_bytes(12)
        
        # Create AES-GCM cipher
        aesgcm = AESGCM(self.encryption_key)
        
        # Prepare associated data
        aad = associated_data.encode('utf-8') if associated_data and isinstance(associated_data, str) else associated_data
        
        # Encrypt
        ciphertext = aesgcm.encrypt(nonce, plaintext_bytes, aad)
        
        # Combine nonce and ciphertext, then base64 encode
        encrypted_data = nonce + ciphertext
        encrypted_b64 = base64.urlsafe_b64encode(encrypted_data).decode('utf-8')
        
        return encrypted_b64
    
    def decrypt(self, encrypted_data, associated_data=None):
        """
        Decrypt encrypted data using AES-256-GCM
        
        Args:
            encrypted_data: Base64-encoded encrypted data with nonce
            associated_data: Optional associated data (must match encryption)
        
        Returns:
            str: Decrypted plaintext string
        """
        if encrypted_data is None:
            return None
        
        try:
            # Decode from base64
            encrypted_bytes = base64.urlsafe_b64decode(encrypted_data.encode('utf-8'))
            
            # Extract nonce (first 12 bytes) and ciphertext (rest)
            nonce = encrypted_bytes[:12]
            ciphertext = encrypted_bytes[12:]
            
            # Create AES-GCM cipher
            aesgcm = AESGCM(self.encryption_key)
            
            # Prepare associated data
            aad = associated_data.encode('utf-8') if associated_data and isinstance(associated_data, str) else associated_data
            
            # Decrypt
            plaintext_bytes = aesgcm.decrypt(nonce, ciphertext, aad)
            
            # Convert to string
            return plaintext_bytes.decode('utf-8')
        
        except Exception as e:
            # If decryption fails, return None or raise exception
            print(f"⚠️ Decryption error: {e}")
            raise ValueError(f"Failed to decrypt data: {e}")
    
    def encrypt_dict(self, data_dict, fields_to_encrypt):
        """
        Encrypt specific fields in a dictionary
        
        Args:
            data_dict: Dictionary containing data
            fields_to_encrypt: List of field names to encrypt
        
        Returns:
            dict: Dictionary with encrypted fields
        """
        encrypted_dict = data_dict.copy()
        for field in fields_to_encrypt:
            if field in encrypted_dict and encrypted_dict[field]:
                encrypted_dict[field] = self.encrypt(encrypted_dict[field])
        return encrypted_dict
    
    def decrypt_dict(self, data_dict, fields_to_decrypt):
        """
        Decrypt specific fields in a dictionary
        
        Args:
            data_dict: Dictionary containing encrypted data
            fields_to_decrypt: List of field names to decrypt
        
        Returns:
            dict: Dictionary with decrypted fields
        """
        decrypted_dict = data_dict.copy()
        for field in fields_to_decrypt:
            if field in decrypted_dict and decrypted_dict[field]:
                try:
                    decrypted_dict[field] = self.decrypt(decrypted_dict[field])
                except Exception as e:
                    print(f"⚠️ Error decrypting field {field}: {e}")
                    decrypted_dict[field] = None
        return decrypted_dict

# Global encryption manager instance
_encryption_manager = None

def get_encryption_manager():
    """Get or create global encryption manager instance"""
    global _encryption_manager
    if _encryption_manager is None:
        _encryption_manager = EncryptionManager()
    return _encryption_manager

def encrypt_data(data, associated_data=None):
    """Convenience function to encrypt data"""
    return get_encryption_manager().encrypt(data, associated_data)

def decrypt_data(encrypted_data, associated_data=None):
    """Convenience function to decrypt data"""
    return get_encryption_manager().decrypt(encrypted_data, associated_data)

