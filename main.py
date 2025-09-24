"""
SharePoint Permissions Analyzer
Main application entry point
"""

import logging
import sys
import os
from datetime import datetime
from typing import Optional
import argparse

from auth import M365Authenticator, get_credentials_interactive
from sharepoint_client import SharePointClient
from html_exporter import HTMLExporter

# Configure logging
def setup_logging(log_level: str = "INFO", log_file: Optional[str] = None):
    """Setup logging configuration"""
    log_format = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    handlers = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file))
    
    logging.basicConfig(
        level=getattr(logging, log_level.upper()),
        format=log_format,
        handlers=handlers
    )

def print_banner():
    """Print application banner"""
    banner = """
╔══════════════════════════════════════════════════════════════════════════════╗
║                    SharePoint Permissions Analyzer                          ║
║                                                                              ║
║  A comprehensive tool to analyze Microsoft 365 SharePoint permissions       ║
║  and shared links across all sites in your tenant.                         ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
    """
    print(banner)

def print_progress(current: int, total: int, message: str):
    """Print progress information"""
    percentage = (current / total) * 100 if total > 0 else 0
    bar_length = 50
    filled_length = int(bar_length * current // total) if total > 0 else 0
    bar = '█' * filled_length + '-' * (bar_length - filled_length)
    
    print(f'\r{message}: |{bar}| {percentage:.1f}% ({current}/{total})', end='', flush=True)
    
    if current == total:
        print()  # New line when complete

def main():
    """Main application function"""
    parser = argparse.ArgumentParser(description="SharePoint Permissions Analyzer")
    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"],
                       help="Set logging level")
    parser.add_argument("--log-file", help="Log to file instead of console")
    parser.add_argument("--output", default="sharepoint_analysis_report.html",
                       help="Output HTML file name")
    parser.add_argument("--tenant-id", help="Microsoft 365 tenant ID (optional)")
    parser.add_argument("--client-id", help="Azure AD client ID (optional)")
    parser.add_argument("--client-secret", help="Azure AD client secret (optional)")
    parser.add_argument("--username", help="Username for authentication (optional)")
    parser.add_argument("--password", help="Password for authentication (optional)")
    parser.add_argument("--setup", action="store_true", help="Run Azure AD app setup")
    parser.add_argument("--delete-creds", action="store_true", help="Delete stored credentials")
    parser.add_argument("--update-creds", action="store_true", help="Update stored credentials")
    
    args = parser.parse_args()
    
    # Setup logging
    setup_logging(args.log_level, args.log_file)
    logger = logging.getLogger(__name__)
    
    print_banner()
    
    # Check if setup is requested
    if args.setup:
        from setup_azure_app import AzureAppSetup
        setup = AzureAppSetup()
        config = setup.setup_interactive()
        if config:
            print("\n✅ Setup complete! You can now run the analyzer with:")
            print("python main.py")
        return 0
    
    # Check if credential management is requested
    if args.delete_creds:
        from credential_manager import CredentialManager
        credential_manager = CredentialManager()
        if credential_manager.has_credentials():
            confirm = input("Are you sure you want to delete stored credentials? (y/N): ").strip().lower()
            if confirm == 'y':
                credential_manager.delete_credentials()
                print("✅ Stored credentials deleted")
            else:
                print("Operation cancelled")
        else:
            print("No stored credentials found")
        return 0
    
    if args.update_creds:
        from credential_manager import CredentialManager
        credential_manager = CredentialManager()
        if credential_manager.has_credentials():
            print("Update stored credentials:")
            client_id = input("New Client ID (or press Enter to keep current): ").strip()
            tenant_id = input("New Tenant ID (or press Enter to keep current): ").strip()
            client_secret = input("New Client Secret (or press Enter to keep current): ").strip()
            
            success = credential_manager.update_credentials(
                client_id=client_id if client_id else None,
                tenant_id=tenant_id if tenant_id else None,
                client_secret=client_secret if client_secret else None
            )
            
            if success:
                print("✅ Credentials updated successfully")
            else:
                print("❌ Failed to update credentials")
        else:
            print("No stored credentials found. Run setup first.")
        return 0
    
    try:
        # Try to load stored credentials if not provided
        if not (args.client_id and args.tenant_id and args.client_secret):
            from credential_manager import CredentialManager
            credential_manager = CredentialManager()
            
            if credential_manager.has_credentials():
                print("Loading stored Azure AD credentials...")
                stored_creds = credential_manager.load_credentials()
                if stored_creds:
                    args.client_id = args.client_id or stored_creds.get("client_id")
                    args.tenant_id = args.tenant_id or stored_creds.get("tenant_id")
                    args.client_secret = args.client_secret or stored_creds.get("client_secret")
                    print("✅ Using stored credentials")
                else:
                    print("❌ Failed to load stored credentials. Please run setup again.")
                    return 1
            else:
                print("No stored credentials found. Please run setup first:")
                print("python main.py --setup")
                return 1
        
        # Initialize authenticator
        logger.info("Initializing Microsoft 365 authenticator...")
        authenticator = M365Authenticator(
            tenant_id=args.tenant_id,
            client_id=args.client_id,
            client_secret=args.client_secret
        )
        
        # Authenticate
        logger.info("Starting authentication process...")
        if args.username and args.password:
            print("Authenticating with provided credentials...")
            auth_success = authenticator.authenticate_with_credentials(args.username, args.password)
        else:
            print("Starting interactive authentication...")
            auth_success = authenticator.authenticate_interactive()
        
        if not auth_success:
            logger.error("Authentication failed. Please check your credentials and try again.")
            return 1
        
        print("✅ Authentication successful!")
        
        # Initialize SharePoint client
        logger.info("Initializing SharePoint client...")
        sp_client = SharePointClient(authenticator)
        
        # Discover sites
        print("\n🔍 Discovering SharePoint sites...")
        sites = sp_client.discover_all_sites()
        
        if not sites:
            logger.warning("No SharePoint sites found. This might indicate:")
            logger.warning("1. Insufficient permissions to access sites")
            logger.warning("2. No sites exist in the tenant")
            logger.warning("3. Sites are not accessible via the current authentication method")
            return 1
        
        print(f"✅ Found {len(sites)} SharePoint sites")
        
        # Analyze all sites
        print(f"\n📊 Analyzing permissions and shared links for {len(sites)} sites...")
        print("This may take several minutes depending on the number of sites and libraries...")
        
        all_analyses = []
        for i, site in enumerate(sites, 1):
            site_name = site.get("displayName", f"Site {i}")
            print_progress(i - 1, len(sites), f"Analyzing {site_name}")
            
            try:
                site_id = site.get("id")
                if site_id:
                    analysis = sp_client.analyze_site_permissions(site_id)
                    if analysis:
                        all_analyses.append(analysis)
            except Exception as e:
                logger.error(f"Error analyzing site {site_name}: {str(e)}")
                continue
        
        print_progress(len(sites), len(sites), "Analysis complete")
        
        if not all_analyses:
            logger.error("No sites could be analyzed. Please check your permissions and try again.")
            return 1
        
        # Calculate summary statistics
        total_sites = len(all_analyses)
        total_libraries = sum(site.get("total_libraries", 0) for site in all_analyses)
        total_shared_links = sum(site.get("total_shared_links", 0) for site in all_analyses)
        total_permissions = sum(site.get("total_permissions", 0) for site in all_analyses)
        
        print(f"\n📈 Analysis Summary:")
        print(f"   • Sites analyzed: {total_sites}")
        print(f"   • Total libraries: {total_libraries}")
        print(f"   • Shared links found: {total_shared_links}")
        print(f"   • Total permissions: {total_permissions}")
        
        # Export to HTML
        print(f"\n📄 Generating HTML report...")
        exporter = HTMLExporter()
        
        try:
            output_file = exporter.export_to_html(all_analyses, args.output)
            print(f"✅ HTML report generated successfully: {output_file}")
            
            # Get absolute path for display
            abs_path = os.path.abspath(output_file)
            print(f"📁 Full path: {abs_path}")
            
        except Exception as e:
            logger.error(f"Error generating HTML report: {str(e)}")
            return 1
        
        # Logout
        authenticator.logout()
        
        print(f"\n🎉 Analysis complete! Report saved to: {output_file}")
        print("You can open the HTML file in your web browser to view the detailed report.")
        
        return 0
        
    except KeyboardInterrupt:
        print("\n\n⚠️  Analysis interrupted by user")
        logger.info("Analysis interrupted by user")
        return 1
        
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        print(f"\n❌ An unexpected error occurred: {str(e)}")
        print("Please check the logs for more details.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
