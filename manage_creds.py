"""
Credential Management Utility
Simple script to manage Azure AD credentials
"""

import sys
from credential_manager import CredentialManager

def main():
    """Main credential management function"""
    print("SharePoint Permissions Analyzer - Credential Manager")
    print("=" * 50)
    print()
    
    manager = CredentialManager()
    
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python manage_creds.py status    - Check credential status")
        print("  python manage_creds.py delete    - Delete stored credentials")
        print("  python manage_creds.py update    - Update stored credentials")
        print("  python manage_creds.py test      - Test credential loading")
        return
    
    command = sys.argv[1].lower()
    
    if command == "status":
        if manager.has_credentials():
            print("✅ Credentials are stored securely")
            try:
                creds = manager.load_credentials()
                if creds:
                    print(f"   Client ID: {creds['client_id'][:8]}...")
                    print(f"   Tenant ID: {creds['tenant_id']}")
                    print(f"   Version: {creds.get('version', 'Unknown')}")
                else:
                    print("   ⚠️  Credentials exist but cannot be decrypted")
            except Exception as e:
                print(f"   ❌ Error loading credentials: {str(e)}")
        else:
            print("❌ No credentials stored")
            print("   Run 'python main.py --setup' to configure credentials")
    
    elif command == "delete":
        if manager.has_credentials():
            confirm = input("Are you sure you want to delete stored credentials? (y/N): ").strip().lower()
            if confirm == 'y':
                manager.delete_credentials()
                print("✅ Credentials deleted successfully")
            else:
                print("Operation cancelled")
        else:
            print("No credentials to delete")
    
    elif command == "update":
        if manager.has_credentials():
            print("Update stored credentials (press Enter to keep current value):")
            client_id = input("Client ID: ").strip()
            tenant_id = input("Tenant ID: ").strip()
            client_secret = input("Client Secret: ").strip()
            
            success = manager.update_credentials(
                client_id=client_id if client_id else None,
                tenant_id=tenant_id if tenant_id else None,
                client_secret=client_secret if client_secret else None
            )
            
            if success:
                print("✅ Credentials updated successfully")
            else:
                print("❌ Failed to update credentials")
        else:
            print("No credentials to update. Run 'python main.py --setup' first.")
    
    elif command == "test":
        if manager.has_credentials():
            print("Testing credential loading...")
            try:
                creds = manager.load_credentials()
                if creds:
                    print("✅ Credentials loaded successfully")
                    print(f"   Client ID: {creds['client_id']}")
                    print(f"   Tenant ID: {creds['tenant_id']}")
                    print(f"   Client Secret: {'*' * len(creds['client_secret'])}")
                else:
                    print("❌ Failed to load credentials")
            except Exception as e:
                print(f"❌ Error: {str(e)}")
        else:
            print("No credentials to test")
    
    else:
        print(f"Unknown command: {command}")
        print("Available commands: status, delete, update, test")

if __name__ == "__main__":
    main()

