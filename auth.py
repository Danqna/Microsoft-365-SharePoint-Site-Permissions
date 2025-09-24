"""
Microsoft 365 Authentication Module
Handles authentication with Microsoft 365 using MSAL
"""

import msal
import getpass
import json
from typing import Optional, Dict, Any
import logging

logger = logging.getLogger(__name__)

class M365Authenticator:
    """Handles Microsoft 365 authentication using MSAL"""
    
    def __init__(self, tenant_id: Optional[str] = None, client_id: Optional[str] = None, client_secret: Optional[str] = None):
        """
        Initialize the authenticator
        
        Args:
            tenant_id: Optional tenant ID. If not provided, will use common endpoint
            client_id: Optional custom client ID
            client_secret: Optional custom client secret
        """
        self.tenant_id = tenant_id or "common"
        self.client_id = client_id or "d3590ed6-52b3-4102-aeff-aad2292ab01c"  # Default Microsoft Office client
        self.client_secret = client_secret
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        # Set scopes based on authentication type
        if client_secret:
            # For client credentials flow, use .default scope
            self.scopes = ["https://graph.microsoft.com/.default"]
        else:
            # For delegated flows, use specific scopes
            self.scopes = [
                "https://graph.microsoft.com/Sites.Read.All",
                "https://graph.microsoft.com/User.Read"
            ]
        self.app = None
        self.access_token = None
        self.account = None
        
    def authenticate_interactive(self) -> bool:
        """
        Perform interactive authentication
        
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            # Create MSAL app
            if self.client_secret:
                # Use confidential client for client secret authentication
                self.app = msal.ConfidentialClientApplication(
                    client_id=self.client_id,
                    client_credential=self.client_secret,
                    authority=self.authority
                )
            else:
                # Use public client for interactive authentication
                self.app = msal.PublicClientApplication(
                    client_id=self.client_id,
                    authority=self.authority
                )
            
            # Try to get token silently first (only for public clients)
            if not self.client_secret:
                accounts = self.app.get_accounts()
                if accounts:
                    result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
                    if result and "access_token" in result:
                        self.access_token = result["access_token"]
                        self.account = accounts[0]
                        logger.info("Successfully authenticated using cached credentials")
                        return True
            
            # Choose authentication method based on client type
            if self.client_secret:
                # Use client credentials flow for confidential clients
                print("Using client credentials authentication...")
                result = self.app.acquire_token_for_client(scopes=self.scopes)
            else:
                # Use device code flow for public clients
                print("Attempting device code authentication...")
                try:
                    flow = self.app.initiate_device_flow(scopes=self.scopes)
                    if "user_code" not in flow:
                        raise ValueError("Fail to create device flow. Err: %s" % json.dumps(flow, indent=2))
                    
                    print(flow["message"])
                    result = self.app.acquire_token_by_device_flow(flow)
                    
                    if "access_token" not in result:
                        # Fall back to interactive authentication
                        print("Device code failed, trying interactive authentication...")
                        result = self.app.acquire_token_interactive(scopes=self.scopes)
                except Exception as e:
                    logger.warning(f"Device code flow failed: {str(e)}")
                    print("Device code failed, trying interactive authentication...")
                    result = self.app.acquire_token_interactive(scopes=self.scopes)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.account = result.get("account")
                logger.info("Successfully authenticated interactively")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def authenticate_with_credentials(self, username: str, password: str) -> bool:
        """
        Authenticate with username and password (less secure, may not work with MFA)
        
        Args:
            username: User's email address
            password: User's password
            
        Returns:
            bool: True if authentication successful, False otherwise
        """
        try:
            # Username/password authentication only works with public clients
            if self.client_secret:
                logger.error("Username/password authentication is not supported with client secret. Use client credentials flow instead.")
                return False
            
            # Use public client for username/password authentication
            self.app = msal.PublicClientApplication(
                client_id=self.client_id,
                authority=self.authority
            )
            
            result = self.app.acquire_token_by_username_password(
                username=username,
                password=password,
                scopes=self.scopes
            )
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                self.account = result.get("account")
                logger.info("Successfully authenticated with credentials")
                return True
            else:
                logger.error(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
                return False
                
        except Exception as e:
            logger.error(f"Authentication error: {str(e)}")
            return False
    
    def get_access_token(self) -> Optional[str]:
        """
        Get the current access token
        
        Returns:
            str: Access token if available, None otherwise
        """
        return self.access_token
    
    def get_headers(self) -> Dict[str, str]:
        """
        Get headers for API requests
        
        Returns:
            dict: Headers with authorization token
        """
        if not self.access_token:
            raise ValueError("Not authenticated. Call authenticate_interactive() or authenticate_with_credentials() first.")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
    
    def refresh_token(self) -> bool:
        """
        Refresh the access token
        
        Returns:
            bool: True if refresh successful, False otherwise
        """
        if not self.app:
            logger.error("Cannot refresh token: not authenticated")
            return False
        
        try:
            if self.client_secret:
                # For confidential clients, use client credentials flow
                result = self.app.acquire_token_for_client(scopes=self.scopes)
            else:
                # For public clients, use silent token acquisition
                if not self.account:
                    logger.error("Cannot refresh token: no account for public client")
                    return False
                result = self.app.acquire_token_silent(self.scopes, account=self.account)
            
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                logger.info("Token refreshed successfully")
                return True
            else:
                logger.error("Failed to refresh token")
                return False
        except Exception as e:
            logger.error(f"Token refresh error: {str(e)}")
            return False
    
    def is_authenticated(self) -> bool:
        """
        Check if currently authenticated
        
        Returns:
            bool: True if authenticated, False otherwise
        """
        return self.access_token is not None
    
    def logout(self):
        """Clear authentication state"""
        if self.app and self.account:
            self.app.remove_account(self.account)
        self.access_token = None
        self.account = None
        logger.info("Logged out successfully")


def get_credentials_interactive() -> tuple[str, str]:
    """
    Get credentials from user input
    
    Returns:
        tuple: (username, password)
    """
    print("Enter your Microsoft 365 credentials:")
    username = input("Username (email): ").strip()
    password = getpass.getpass("Password: ")
    return username, password
