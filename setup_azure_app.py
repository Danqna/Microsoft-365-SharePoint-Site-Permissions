"""
Azure AD App Registration Setup
Automates the creation of an Azure AD app registration for SharePoint access
"""

import requests
import json
import webbrowser
import time
from urllib.parse import urlencode, parse_qs, urlparse
import secrets
import string

class AzureAppSetup:
    """Handles Azure AD app registration setup"""
    
    def __init__(self):
        self.client_id = None
        self.client_secret = None
        self.tenant_id = None
        
    def generate_client_secret(self, length=32):
        """Generate a random client secret"""
        alphabet = string.ascii_letters + string.digits + "!@#$%^&*"
        return ''.join(secrets.choice(alphabet) for _ in range(length))
    
    def create_app_registration_instructions(self):
        """Generate step-by-step instructions for manual app registration"""
        
        print("=" * 80)
        print("AZURE AD APP REGISTRATION SETUP")
        print("=" * 80)
        print()
        print("Since we can't automatically create the app registration, please follow these steps:")
        print()
        print("1. Go to Azure Portal: https://portal.azure.com")
        print("2. Navigate to: Azure Active Directory > App registrations")
        print("3. Click 'New registration'")
        print("4. Fill in the details:")
        print("   - Name: SharePoint Permissions Analyzer")
        print("   - Supported account types: Accounts in this organizational directory only")
        print("   - Redirect URI: http://localhost:8080 (Web)")
        print("5. Click 'Register'")
        print()
        print("6. Copy the 'Application (client) ID' and 'Directory (tenant) ID'")
        print()
        print("7. Go to 'Certificates & secrets' > 'New client secret'")
        print("   - Description: SharePoint Analyzer Secret")
        print("   - Expires: 24 months")
        print("8. Copy the secret VALUE (not the ID)")
        print()
        print("9. Go to 'API permissions' > 'Add a permission'")
        print("   - Select 'Microsoft Graph'")
        print("   - Select 'Application permissions' (NOT Delegated permissions)")
        print("   - Add these permissions:")
        print("     - Sites.Read.All")
        print("     - User.Read.All")
        print("10. Click 'Grant admin consent' (REQUIRED for application permissions)")
        print()
        print("11. Save the following information:")
        print("   - Client ID: [Your Application ID]")
        print("   - Client Secret: [Your Secret Value]")
        print("   - Tenant ID: [Your Directory ID]")
        print()
        print("=" * 80)
        
        return {
            "client_id": input("Enter your Client ID: ").strip(),
            "client_secret": input("Enter your Client Secret: ").strip(),
            "tenant_id": input("Enter your Tenant ID: ").strip()
        }
    
    def save_config(self, config):
        """Save the configuration securely"""
        from credential_manager import CredentialManager
        
        credential_manager = CredentialManager()
        
        # Save credentials securely
        success = credential_manager.save_credentials(
            client_id=config["client_id"],
            tenant_id=config["tenant_id"],
            client_secret=config["client_secret"]
        )
        
        if success:
            print("✅ Credentials saved securely!")
            print("Your Azure AD credentials are now encrypted and stored locally.")
            print("You won't need to enter them again unless you change your password.")
        else:
            print("❌ Failed to save credentials securely")
        
        return config
    
    def load_config(self):
        """Load configuration from secure storage"""
        from credential_manager import CredentialManager
        
        credential_manager = CredentialManager()
        
        if credential_manager.has_credentials():
            print("Found existing Azure AD configuration.")
            credentials = credential_manager.load_credentials()
            if credentials:
                return credentials
            else:
                print("Failed to load existing credentials. You may need to run setup again.")
        
        return None
    
    def setup_interactive(self):
        """Interactive setup process"""
        print("SharePoint Permissions Analyzer - Azure AD Setup")
        print("=" * 50)
        print()
        
        # Try to load existing config
        config = self.load_config()
        
        if not config:
            print("No existing configuration found.")
            print("You need to create an Azure AD app registration.")
            print()
            
            choice = input("Do you want to see setup instructions? (y/n): ").strip().lower()
            if choice == 'y':
                config = self.create_app_registration_instructions()
                self.save_config(config)
            else:
                print("Setup cancelled.")
                return None
        
        return config

def main():
    """Main setup function"""
    setup = AzureAppSetup()
    config = setup.setup_interactive()
    
    if config:
        print("\n✅ Azure AD configuration complete!")
        print("You can now run the SharePoint analyzer.")
        print("\nTo run the analyzer:")
        print("python main.py --client-id {} --tenant-id {} --client-secret {}".format(
            config["client_id"], config["tenant_id"], config["client_secret"]
        ))
    else:
        print("\n❌ Setup incomplete. Please run this script again when ready.")

if __name__ == "__main__":
    main()
