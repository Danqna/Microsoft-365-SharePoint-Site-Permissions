"""
Configuration settings for SharePoint Permissions Analyzer
"""

import os
from typing import Optional

class Config:
    """Configuration class for the application"""
    
    # Microsoft Graph API settings
    GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0"
    # Scopes are set dynamically based on authentication type
    # For client credentials: ["https://graph.microsoft.com/.default"]
    # For delegated flows: ["https://graph.microsoft.com/Sites.Read.All", "https://graph.microsoft.com/User.Read"]
    
    # Authentication settings
    CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"  # Microsoft Graph PowerShell client
    TENANT_ID = os.getenv("TENANT_ID", "common")  # Can be overridden with environment variable
    
    # Application settings
    DEFAULT_OUTPUT_FILE = "sharepoint_analysis_report.html"
    LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
    LOG_FILE = os.getenv("LOG_FILE")
    
    # Request settings
    REQUEST_TIMEOUT = 30  # seconds
    MAX_RETRIES = 3
    RETRY_DELAY = 1  # seconds
    
    # Rate limiting
    RATE_LIMIT_DELAY = 1  # seconds between requests
    MAX_REQUESTS_PER_MINUTE = 60
    
    # HTML export settings
    HTML_TEMPLATE_FILE = None  # Use built-in template
    INCLUDE_SITE_DETAILS = True
    INCLUDE_LIBRARY_DETAILS = True
    INCLUDE_SHARED_LINKS = True
    INCLUDE_PERMISSIONS = True
    
    # Security settings
    MASK_SENSITIVE_DATA = True
    EXCLUDE_SYSTEM_SITES = True
    
    @classmethod
    def get_tenant_id(cls) -> str:
        """Get tenant ID from environment or default"""
        return cls.TENANT_ID
    
    @classmethod
    def get_client_id(cls) -> str:
        """Get client ID from environment or default"""
        return os.getenv("CLIENT_ID", cls.CLIENT_ID)
    
    @classmethod
    def get_output_file(cls) -> str:
        """Get output file from environment or default"""
        return os.getenv("OUTPUT_FILE", cls.DEFAULT_OUTPUT_FILE)
    
    @classmethod
    def is_debug_mode(cls) -> bool:
        """Check if debug mode is enabled"""
        return cls.LOG_LEVEL.upper() == "DEBUG"
    
    @classmethod
    def should_exclude_system_sites(cls) -> bool:
        """Check if system sites should be excluded"""
        return cls.EXCLUDE_SYSTEM_SITES
    
    @classmethod
    def get_rate_limit_delay(cls) -> float:
        """Get rate limit delay in seconds"""
        return float(os.getenv("RATE_LIMIT_DELAY", cls.RATE_LIMIT_DELAY))
