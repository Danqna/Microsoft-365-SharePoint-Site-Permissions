"""
Secure Credential Manager
Handles encrypted storage and retrieval of Azure AD credentials
"""

import json
import base64
import os
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import getpass
import logging

logger = logging.getLogger(__name__)

class CredentialManager:
    """Manages secure storage of Azure AD credentials"""
    
    def __init__(self, config_file="azure_credentials.json"):
        self.config_file = config_file
        self.key_file = "credentials.key"
        
    def _get_or_create_key(self, password: str = None) -> bytes:
        """Get or create encryption key"""
        if os.path.exists(self.key_file):
            with open(self.key_file, 'rb') as f:
                return f.read()
        
        # Create new key
        if not password:
            password = getpass.getpass("Enter password to encrypt credentials (or press Enter for auto-generated): ")
            if not password:
                # Generate random password
                password = base64.urlsafe_b64encode(os.urandom(32)).decode()
                print(f"Auto-generated password: {password}")
                print("IMPORTANT: Save this password! You'll need it to decrypt credentials.")
                input("Press Enter to continue...")
        
        # Derive key from password
        salt = os.urandom(16)
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        
        # Save key and salt
        with open(self.key_file, 'wb') as f:
            f.write(salt + key)
        
        return key
    
    def _load_key(self, password: str = None) -> bytes:
        """Load encryption key from file"""
        if not os.path.exists(self.key_file):
            raise FileNotFoundError("No credentials file found. Run setup first.")
        
        with open(self.key_file, 'rb') as f:
            data = f.read()
        
        salt = data[:16]
        stored_key = data[16:]
        
        if not password:
            password = getpass.getpass("Enter password to decrypt credentials: ")
        
        # Derive key from password
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        
        if key != stored_key:
            raise ValueError("Invalid password")
        
        return key
    
    def save_credentials(self, client_id: str, tenant_id: str, client_secret: str, password: str = None) -> bool:
        """Save credentials securely"""
        try:
            # Get or create encryption key
            key = self._get_or_create_key(password)
            fernet = Fernet(key)
            
            # Prepare credentials
            credentials = {
                "client_id": client_id,
                "tenant_id": tenant_id,
                "client_secret": client_secret,
                "version": "1.0"
            }
            
            # Encrypt credentials
            encrypted_data = fernet.encrypt(json.dumps(credentials).encode())
            
            # Save to file
            with open(self.config_file, 'wb') as f:
                f.write(encrypted_data)
            
            logger.info("Credentials saved securely")
            return True
            
        except Exception as e:
            logger.error(f"Failed to save credentials: {str(e)}")
            return False
    
    def load_credentials(self, password: str = None) -> dict:
        """Load credentials securely"""
        try:
            if not os.path.exists(self.config_file):
                return None
            
            # Load encryption key
            key = self._load_key(password)
            fernet = Fernet(key)
            
            # Load and decrypt credentials
            with open(self.config_file, 'rb') as f:
                encrypted_data = f.read()
            
            decrypted_data = fernet.decrypt(encrypted_data)
            credentials = json.loads(decrypted_data.decode())
            
            logger.info("Credentials loaded successfully")
            return credentials
            
        except Exception as e:
            logger.error(f"Failed to load credentials: {str(e)}")
            return None
    
    def has_credentials(self) -> bool:
        """Check if credentials exist"""
        return os.path.exists(self.config_file) and os.path.exists(self.key_file)
    
    def delete_credentials(self):
        """Delete stored credentials"""
        try:
            if os.path.exists(self.config_file):
                os.remove(self.config_file)
            if os.path.exists(self.key_file):
                os.remove(self.key_file)
            logger.info("Credentials deleted")
        except Exception as e:
            logger.error(f"Failed to delete credentials: {str(e)}")
    
    def update_credentials(self, client_id: str = None, tenant_id: str = None, client_secret: str = None, password: str = None) -> bool:
        """Update existing credentials"""
        try:
            # Load existing credentials
            existing = self.load_credentials(password)
            if not existing:
                return False
            
            # Update only provided fields
            if client_id:
                existing["client_id"] = client_id
            if tenant_id:
                existing["tenant_id"] = tenant_id
            if client_secret:
                existing["client_secret"] = client_secret
            
            # Save updated credentials
            return self.save_credentials(
                existing["client_id"],
                existing["tenant_id"],
                existing["client_secret"],
                password
            )
            
        except Exception as e:
            logger.error(f"Failed to update credentials: {str(e)}")
            return False

def main():
    """Test the credential manager"""
    manager = CredentialManager()
    
    print("Credential Manager Test")
    print("======================")
    
    # Test saving credentials
    print("\n1. Testing credential storage...")
    success = manager.save_credentials(
        client_id="test-client-id",
        tenant_id="test-tenant-id",
        client_secret="test-client-secret"
    )
    print(f"Save result: {success}")
    
    # Test loading credentials
    print("\n2. Testing credential loading...")
    credentials = manager.load_credentials()
    if credentials:
        print(f"Loaded: {credentials}")
    else:
        print("Failed to load credentials")
    
    # Test credential existence
    print(f"\n3. Credentials exist: {manager.has_credentials()}")
    
    # Clean up
    print("\n4. Cleaning up...")
    manager.delete_credentials()

if __name__ == "__main__":
    main()

