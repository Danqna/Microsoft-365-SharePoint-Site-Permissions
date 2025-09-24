# SharePoint Permissions Analyzer

A comprehensive Python application to analyze Microsoft 365 SharePoint permissions and shared links across all sites in your tenant.

## Features

- 🔐 **Modern Authentication**: Secure authentication with Microsoft 365 using MSAL
- 🔍 **Site Discovery**: Automatically discover all SharePoint sites in your tenant
- 📚 **Library Analysis**: Analyze document libraries and their shared links
- 🔑 **Permissions Audit**: Extract detailed permissions for each library
- 📊 **HTML Reports**: Export results to beautifully formatted HTML reports
- 🚀 **Easy to Use**: Simple command-line interface with progress tracking
- ⚡ **Efficient**: Optimized API calls with rate limiting and error handling

## Installation

### Prerequisites
- Python 3.8 or higher
- Microsoft 365 tenant with SharePoint
- Appropriate permissions to read site collections and permissions
- Internet connection for authentication

### Quick Start

1. **Download** this repository

2. **Double-click** `run_analyzer.bat` (everything is automated!)
   
   That's it! The script will automatically:
   - ✅ Check Python installation
   - ✅ Create/activate virtual environment
   - ✅ Install dependencies
   - ✅ Check for Azure AD credentials
   - ✅ Guide you through Azure AD setup (first time only)
   - ✅ Run the SharePoint analysis
   - ✅ Generate HTML report

**No command line knowledge required!** Just double-click and follow the prompts.

### Windows Users

**Simple Option (Recommended):**
```cmd
run_analyzer.bat
```

**PowerShell Option (with more options):**
```powershell
.\run_analyzer.ps1
```

Both scripts handle everything automatically:
- ✅ **Python Environment**: Creates/activates virtual environment
- ✅ **Dependencies**: Installs all required packages
- ✅ **Azure AD Setup**: Guides you through app registration (first time)
- ✅ **Credential Storage**: Securely stores your credentials encrypted
- ✅ **Analysis**: Runs the SharePoint permissions analysis
- ✅ **Cleanup**: Properly deactivates virtual environment

**No manual setup required!** Just double-click and follow the prompts.

## Usage

### Basic Usage

```bash
python main.py
```

The application will:
1. Load your stored Azure AD app credentials (encrypted)
2. Connect to SharePoint and discover all sites
3. Analyze permissions and shared links for each site
4. Generate a comprehensive HTML report

### Advanced Options

```bash
python main.py --help
```

Available options:
- `--log-level`: Set logging level (DEBUG, INFO, WARNING, ERROR)
- `--log-file`: Log to file instead of console
- `--output`: Specify output HTML file name
- `--setup`: Run Azure AD app registration setup
- `--update-creds`: Update stored Azure AD credentials
- `--delete-creds`: Delete stored Azure AD credentials

### Examples

```bash
# Basic analysis with stored credentials
python main.py

# Debug mode with custom output file
python main.py --log-level DEBUG --output my_analysis.html

# Run Azure AD app setup (first time only)
python main.py --setup

# Update stored credentials
python main.py --update-creds

# Delete stored credentials
python main.py --delete-creds
```

## Authentication

The application requires an Azure AD app registration to authenticate with Microsoft Graph API. This is necessary because Microsoft's first-party client IDs require special preauthorization.

### Azure AD App Registration Setup

**Option 1: Automated Setup (Recommended)**
```bash
python main.py --setup
```
This will guide you through the manual setup process with step-by-step instructions.

**Option 2: Manual Setup**
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to: Azure Active Directory > App registrations
3. Click "New registration"
4. Fill in:
   - Name: SharePoint Permissions Analyzer
   - Supported account types: Accounts in this organizational directory only
   - Redirect URI: http://localhost:8080 (Web)
5. Click "Register"
6. Copy the Application (client) ID and Directory (tenant) ID
7. Go to "Certificates & secrets" > "New client secret"
8. Copy the secret VALUE
9. Go to "API permissions" > "Add a permission" > "Microsoft Graph" > "Application permissions"
10. Add: `Sites.Read.All` and `User.Read.All`
11. Click "Grant admin consent" (REQUIRED for application permissions)

### Authentication Methods

1. **Client Secret Authentication** (Recommended): Uses your registered app's client ID and secret
2. **Device Code Authentication**: For environments without browser access
3. **Interactive Authentication**: Opens a browser window for secure authentication (fallback)

### Secure Credential Storage

The application stores your Azure AD app credentials securely using:
- **AES encryption** with PBKDF2 key derivation
- **Password-protected** credential files (you set the encryption password)
- **Local storage only** - credentials never leave your machine
- **Automatic loading** - credentials are loaded automatically when you run the analyzer

**Credential Management Commands:**
```bash
# Update stored credentials
python main.py --update-creds

# Delete stored credentials
python main.py --delete-creds
```

### Required Permissions

Your Azure AD app needs the following Microsoft Graph permissions:

**For Client Credentials Flow (Recommended):**
- `Sites.Read.All` - Application permission (not delegated)
- `User.Read.All` - Application permission (not delegated)

**For Delegated Flow:**
- `Sites.Read.All` - Delegated permission
- `User.Read` - Delegated permission

**Important:** When using client credentials flow (with client secret), you need **Application permissions**, not Delegated permissions. 

⚠️ **Common Mistake:** If you accidentally select "Delegated permissions" instead of "Application permissions", you'll get 401 errors. You must use "Application permissions" for client credentials flow.

## Output

The application generates a comprehensive HTML report containing:

### Executive Summary
- Total number of sites discovered
- Total libraries analyzed
- Total shared links found
- Total permissions identified

### Site Details
- Site name, URL, and metadata
- Creation and modification dates
- List of all document libraries

### Library Analysis
- Library name, description, and URL
- Shared links with file details
- Detailed permissions breakdown
- Permission types (Direct, Inherited, Link-based)

### Interactive Features
- Collapsible sections for easy navigation
- Color-coded permission badges
- Responsive design for mobile viewing
- Direct links to SharePoint sites and files

## File Structure

```
├── run_analyzer.bat       # 🚀 MAIN SCRIPT - Run this!
├── run_analyzer.ps1       # PowerShell version with options
├── main.py                # Core application
├── auth.py                # Microsoft 365 authentication
├── sharepoint_client.py   # SharePoint API operations
├── html_exporter.py       # HTML report generation
├── credential_manager.py  # Secure credential storage
├── setup_azure_app.py     # Azure AD setup helper
├── manage_creds.py        # Credential management utility
├── requirements.txt       # Python dependencies
└── README.md              # This file
```

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Ensure you have the correct Azure AD app permissions
   - Verify your client secret hasn't expired
   - Check that admin consent has been granted for application permissions

2. **No Sites Found**
   - Verify you have access to SharePoint sites
   - Check if the tenant has any SharePoint sites
   - Ensure you're using the correct tenant ID

3. **Rate Limiting**
   - The application automatically handles rate limiting
   - Large tenants may take longer to analyze
   - Consider running during off-peak hours

4. **Permission Errors**
   - Ensure your account has `Sites.Read.All` permission
   - Contact your administrator if permissions are insufficient

5. **401 Errors with Valid Authentication**
   - Check that you're using **Application permissions** (not Delegated permissions)
   - Ensure admin consent has been granted
   - Verify the permissions are `Sites.Read.All` and `User.Read.All` (Application type)

### Logging

Enable debug logging for detailed information:
```bash
python main.py --log-level DEBUG --log-file analysis.log
```

## Security Considerations

- Azure AD app credentials are encrypted and stored locally
- Encryption password is never stored (you enter it each time)
- All API calls use HTTPS
- Tokens are automatically refreshed as needed
- Credentials can be deleted at any time using `--delete-creds`

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## License

This project is provided as-is for educational and administrative purposes.
