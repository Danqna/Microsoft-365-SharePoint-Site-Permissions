"""
SharePoint Permissions Diagnostic Tool
Helps diagnose authentication and permission issues
"""

import requests
import json
from auth import M365Authenticator
from credential_manager import CredentialManager
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def test_graph_api_access(authenticator):
    """Test various Graph API endpoints to diagnose issues"""
    
    base_url = "https://graph.microsoft.com/v1.0"
    headers = authenticator.get_headers()
    
    print("🔍 Testing Microsoft Graph API Access")
    print("=" * 50)
    
    # Test 1: Basic user info
    print("\n1. Testing basic user access (/me)...")
    try:
        response = requests.get(f"{base_url}/me", headers=headers)
        if response.status_code == 200:
            user_data = response.json()
            print(f"   ✅ Success: Authenticated as {user_data.get('displayName', 'Unknown')}")
            print(f"   User ID: {user_data.get('id', 'Unknown')}")
        else:
            print(f"   ❌ Failed: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
    
    # Test 2: Organization info
    print("\n2. Testing organization access (/organization)...")
    try:
        response = requests.get(f"{base_url}/organization", headers=headers)
        if response.status_code == 200:
            org_data = response.json()
            print(f"   ✅ Success: Found {len(org_data.get('value', []))} organization(s)")
            for org in org_data.get('value', []):
                print(f"   Organization: {org.get('displayName', 'Unknown')}")
        else:
            print(f"   ❌ Failed: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
    
    # Test 3: Sites root
    print("\n3. Testing SharePoint root site access (/sites/root)...")
    try:
        response = requests.get(f"{base_url}/sites/root", headers=headers)
        if response.status_code == 200:
            site_data = response.json()
            print(f"   ✅ Success: Root site found")
            print(f"   Site ID: {site_data.get('id', 'Unknown')}")
            print(f"   Site Name: {site_data.get('displayName', 'Unknown')}")
            print(f"   Site URL: {site_data.get('webUrl', 'Unknown')}")
        else:
            print(f"   ❌ Failed: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
    
    # Test 4: Sites search
    print("\n4. Testing SharePoint sites search (/sites?search=*)...")
    try:
        response = requests.get(f"{base_url}/sites?search=*", headers=headers)
        if response.status_code == 200:
            sites_data = response.json()
            sites = sites_data.get('value', [])
            print(f"   ✅ Success: Found {len(sites)} sites")
            for site in sites[:5]:  # Show first 5 sites
                print(f"   - {site.get('displayName', 'Unknown')}: {site.get('webUrl', 'Unknown')}")
        else:
            print(f"   ❌ Failed: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
    
    # Test 5: Drives
    print("\n5. Testing drives access (/drives)...")
    try:
        response = requests.get(f"{base_url}/drives", headers=headers)
        if response.status_code == 200:
            drives_data = response.json()
            drives = drives_data.get('value', [])
            print(f"   ✅ Success: Found {len(drives)} drives")
            for drive in drives[:3]:  # Show first 3 drives
                print(f"   - {drive.get('name', 'Unknown')}: {drive.get('webUrl', 'Unknown')}")
        else:
            print(f"   ❌ Failed: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"   ❌ Error: {str(e)}")
    
    # Test 6: Token info
    print("\n6. Analyzing access token...")
    try:
        token = authenticator.get_access_token()
        if token:
            # Decode JWT token (basic info only)
            import base64
            parts = token.split('.')
            if len(parts) >= 2:
                # Add padding if needed
                payload = parts[1]
                payload += '=' * (4 - len(payload) % 4)
                decoded = base64.urlsafe_b64decode(payload)
                token_data = json.loads(decoded)
                
                print(f"   ✅ Token decoded successfully")
                print(f"   Issued at: {token_data.get('iat', 'Unknown')}")
                print(f"   Expires at: {token_data.get('exp', 'Unknown')}")
                print(f"   Audience: {token_data.get('aud', 'Unknown')}")
                print(f"   Scopes: {token_data.get('scp', 'Unknown')}")
                print(f"   Roles: {token_data.get('roles', 'None')}")
        else:
            print(f"   ❌ No access token available")
    except Exception as e:
        print(f"   ❌ Error decoding token: {str(e)}")

def main():
    """Main diagnostic function"""
    print("SharePoint Permissions Diagnostic Tool")
    print("=" * 50)
    
    # Load credentials
    credential_manager = CredentialManager()
    if not credential_manager.has_credentials():
        print("❌ No stored credentials found. Please run setup first:")
        print("python main.py --setup")
        return
    
    print("Loading stored credentials...")
    credentials = credential_manager.load_credentials()
    if not credentials:
        print("❌ Failed to load credentials")
        return
    
    print("✅ Credentials loaded successfully")
    
    # Initialize authenticator
    authenticator = M365Authenticator(
        tenant_id=credentials.get("tenant_id"),
        client_id=credentials.get("client_id"),
        client_secret=credentials.get("client_secret")
    )
    
    # Authenticate
    print("\nAuthenticating...")
    if not authenticator.authenticate_interactive():
        print("❌ Authentication failed")
        return
    
    print("✅ Authentication successful")
    
    # Run diagnostics
    test_graph_api_access(authenticator)
    
    print("\n" + "=" * 50)
    print("Diagnostic complete!")
    print("\nIf you see ❌ errors above, the issue is likely:")
    print("1. Application permissions not granted in Azure AD")
    print("2. Wrong permission type (should be Application, not Delegated)")
    print("3. Missing admin consent")
    print("4. Tenant restrictions on application permissions")

if __name__ == "__main__":
    main()

