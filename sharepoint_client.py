"""
SharePoint Client Module
Handles SharePoint operations including site discovery, library enumeration, and permissions analysis
"""

import requests
import json
import logging
from typing import List, Dict, Any, Optional
from urllib.parse import urljoin
import time

logger = logging.getLogger(__name__)

class SharePointClient:
    """Client for SharePoint operations using Microsoft Graph API"""
    
    def __init__(self, authenticator):
        """
        Initialize SharePoint client
        
        Args:
            authenticator: M365Authenticator instance
        """
        self.authenticator = authenticator
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.sites = []
        
        # Test authentication and permissions
        self._test_authentication()
    
    def _test_authentication(self):
        """Test authentication and basic permissions"""
        logger.info("Testing authentication and permissions...")
        
        # Test basic Graph API access
        test_url = f"{self.base_url}/me"
        response = self._make_request(test_url)
        
        if response:
            logger.info("✅ Basic Graph API access working")
            logger.info(f"Authenticated as: {response.get('displayName', 'Unknown')}")
        else:
            logger.warning("⚠️ Basic Graph API access failed")
        
        # Test SharePoint access
        test_url = f"{self.base_url}/sites/root"
        response = self._make_request(test_url)
        
        if response:
            logger.info("✅ SharePoint root site access working")
        else:
            logger.warning("⚠️ SharePoint root site access failed")
        
    def _make_request(self, url: str, method: str = "GET", data: Optional[Dict] = None) -> Optional[Dict]:
        """
        Make authenticated request to Microsoft Graph API
        
        Args:
            url: API endpoint URL
            method: HTTP method
            data: Request data for POST/PUT requests
            
        Returns:
            dict: Response data or None if failed
        """
        try:
            headers = self.authenticator.get_headers()
            logger.debug(f"Making request to: {url}")
            
            if method.upper() == "GET":
                response = requests.get(url, headers=headers)
            elif method.upper() == "POST":
                response = requests.post(url, headers=headers, json=data)
            elif method.upper() == "PUT":
                response = requests.put(url, headers=headers, json=data)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")
            
            if response.status_code == 401:
                # Token might be expired, try to refresh
                logger.warning("Token expired, attempting to refresh...")
                if self.authenticator.refresh_token():
                    headers = self.authenticator.get_headers()
                    if method.upper() == "GET":
                        response = requests.get(url, headers=headers)
                    elif method.upper() == "POST":
                        response = requests.post(url, headers=headers, json=data)
                    elif method.upper() == "PUT":
                        response = requests.put(url, headers=headers, json=data)
                else:
                    logger.error("Failed to refresh token")
                    return None
            
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 429:
                # Rate limited, wait and retry
                retry_after = int(response.headers.get('Retry-After', 60))
                logger.warning(f"Rate limited, waiting {retry_after} seconds...")
                time.sleep(retry_after)
                return self._make_request(url, method, data)
            else:
                logger.error(f"API request failed: {response.status_code} - {response.text}")
                logger.error(f"Request URL: {url}")
                logger.error(f"Request headers: {headers}")
                return None
                
        except Exception as e:
            logger.error(f"Request error: {str(e)}")
            return None
    
    def discover_all_sites(self) -> List[Dict[str, Any]]:
        """
        Discover all SharePoint sites in the tenant
        
        Returns:
            list: List of site information dictionaries
        """
        logger.info("Discovering SharePoint sites...")
        sites = []
        
        # Try different approaches to discover sites
        discovery_methods = [
            self._discover_sites_via_search,
            self._discover_sites_via_root,
            self._discover_sites_via_drives
        ]
        
        for method in discovery_methods:
            try:
                method_sites = method()
                if method_sites:
                    sites.extend(method_sites)
                    logger.info(f"Found {len(method_sites)} sites using {method.__name__}")
                    break  # If we found sites with one method, use those
            except Exception as e:
                logger.warning(f"Method {method.__name__} failed: {str(e)}")
                continue
        
        # Remove duplicates
        unique_sites = []
        seen_ids = set()
        for site in sites:
            site_id = site.get("id")
            if site_id and site_id not in seen_ids:
                unique_sites.append(site)
                seen_ids.add(site_id)
        
        logger.info(f"Discovered {len(unique_sites)} SharePoint sites")
        self.sites = unique_sites
        return unique_sites
    
    def _discover_sites_via_search(self) -> List[Dict[str, Any]]:
        """Try to discover sites using search endpoint"""
        sites = []
        
        # Try the search endpoint
        url = f"{self.base_url}/sites?search=*"
        data = self._make_request(url)
        
        if data and "value" in data:
            sites.extend(data["value"])
            
            # Handle pagination
            while "@odata.nextLink" in data:
                logger.info("Fetching next page of sites...")
                data = self._make_request(data["@odata.nextLink"])
                if data and "value" in data:
                    sites.extend(data["value"])
        
        return sites
    
    def _discover_sites_via_root(self) -> List[Dict[str, Any]]:
        """Try to discover sites starting from root site"""
        sites = []
        
        try:
            # Get the root site
            root_site_url = f"{self.base_url}/sites/root"
            root_data = self._make_request(root_site_url)
            if root_data:
                sites.append(root_data)
                logger.info("Found root site")
        except Exception as e:
            logger.warning(f"Could not fetch root site: {str(e)}")
        
        return sites
    
    def _discover_sites_via_drives(self) -> List[Dict[str, Any]]:
        """Try to discover sites by getting drives (this might work with application permissions)"""
        sites = []
        
        try:
            # Try to get drives directly
            drives_url = f"{self.base_url}/drives"
            drives_data = self._make_request(drives_url)
            
            if drives_data and "value" in drives_data:
                # Extract site information from drives
                for drive in drives_data["value"]:
                    if "parentReference" in drive and "siteId" in drive["parentReference"]:
                        site_id = drive["parentReference"]["siteId"]
                        site_url = f"{self.base_url}/sites/{site_id}"
                        site_data = self._make_request(site_url)
                        if site_data:
                            sites.append(site_data)
        except Exception as e:
            logger.warning(f"Could not discover sites via drives: {str(e)}")
        
        return sites
    
    def get_site_libraries(self, site_id: str) -> List[Dict[str, Any]]:
        """
        Get all document libraries for a specific site
        
        Args:
            site_id: SharePoint site ID
            
        Returns:
            list: List of library information dictionaries
        """
        logger.info(f"Getting libraries for site: {site_id}")
        libraries = []
        
        try:
            # Get drives (document libraries)
            url = f"{self.base_url}/sites/{site_id}/drives"
            data = self._make_request(url)
            
            if data and "value" in data:
                for drive in data["value"]:
                    library_info = {
                        "id": drive.get("id"),
                        "name": drive.get("name"),
                        "description": drive.get("description", ""),
                        "webUrl": drive.get("webUrl"),
                        "driveType": drive.get("driveType"),
                        "createdDateTime": drive.get("createdDateTime"),
                        "lastModifiedDateTime": drive.get("lastModifiedDateTime"),
                        "owner": drive.get("owner", {}),
                        "quota": drive.get("quota", {}),
                        "shared_links": [],
                        "permissions": []
                    }
                    libraries.append(library_info)
            
            logger.info(f"Found {len(libraries)} libraries in site {site_id}")
            return libraries
            
        except Exception as e:
            logger.error(f"Error getting libraries for site {site_id}: {str(e)}")
            return []
    
    def get_library_shared_links(self, site_id: str, drive_id: str) -> List[Dict[str, Any]]:
        """
        Get shared links for a specific library
        
        Args:
            site_id: SharePoint site ID
            drive_id: Drive/library ID
            
        Returns:
            list: List of shared link information
        """
        logger.info(f"Getting shared links for library: {drive_id}")
        shared_links = []
        
        try:
            # Get shared items
            url = f"{self.base_url}/sites/{site_id}/drives/{drive_id}/shared"
            data = self._make_request(url)
            
            if data and "value" in data:
                for item in data["value"]:
                    link_info = {
                        "id": item.get("id"),
                        "name": item.get("name"),
                        "webUrl": item.get("webUrl"),
                        "downloadUrl": item.get("@microsoft.graph.downloadUrl"),
                        "createdDateTime": item.get("createdDateTime"),
                        "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                        "size": item.get("size"),
                        "createdBy": item.get("createdBy", {}),
                        "lastModifiedBy": item.get("lastModifiedBy", {}),
                        "shared": item.get("shared", {})
                    }
                    shared_links.append(link_info)
            
            logger.info(f"Found {len(shared_links)} shared items in library {drive_id}")
            return shared_links
            
        except Exception as e:
            logger.error(f"Error getting shared links for library {drive_id}: {str(e)}")
            return []
    
    def get_library_permissions(self, site_id: str, drive_id: str) -> List[Dict[str, Any]]:
        """
        Get permissions for a specific library
        
        Args:
            site_id: SharePoint site ID
            drive_id: Drive/library ID
            
        Returns:
            list: List of permission information
        """
        logger.info(f"Getting permissions for library: {drive_id}")
        permissions = []
        
        try:
            # Get permissions for the drive
            url = f"{self.base_url}/sites/{site_id}/drives/{drive_id}/permissions"
            data = self._make_request(url)
            
            if data and "value" in data:
                for permission in data["value"]:
                    perm_info = {
                        "id": permission.get("id"),
                        "roles": permission.get("roles", []),
                        "grantedTo": permission.get("grantedTo", {}),
                        "grantedToIdentities": permission.get("grantedToIdentities", []),
                        "link": permission.get("link", {}),
                        "inheritedFrom": permission.get("inheritedFrom", {}),
                        "expirationDateTime": permission.get("expirationDateTime")
                    }
                    permissions.append(perm_info)
            
            logger.info(f"Found {len(permissions)} permissions for library {drive_id}")
            return permissions
            
        except Exception as e:
            logger.error(f"Error getting permissions for library {drive_id}: {str(e)}")
            return []
    
    def analyze_site_permissions(self, site_id: str) -> Dict[str, Any]:
        """
        Analyze all permissions and shared links for a site
        
        Args:
            site_id: SharePoint site ID
            
        Returns:
            dict: Complete analysis data for the site
        """
        logger.info(f"Analyzing permissions for site: {site_id}")
        
        # Get site information
        site_url = f"{self.base_url}/sites/{site_id}"
        site_data = self._make_request(site_url)
        
        if not site_data:
            logger.error(f"Could not get site data for {site_id}")
            return {}
        
        # Get libraries
        libraries = self.get_site_libraries(site_id)
        
        # Analyze each library
        for library in libraries:
            drive_id = library["id"]
            
            # Get shared links
            library["shared_links"] = self.get_library_shared_links(site_id, drive_id)
            
            # Get permissions
            library["permissions"] = self.get_library_permissions(site_id, drive_id)
        
        analysis = {
            "site_info": site_data,
            "libraries": libraries,
            "total_libraries": len(libraries),
            "total_shared_links": sum(len(lib["shared_links"]) for lib in libraries),
            "total_permissions": sum(len(lib["permissions"]) for lib in libraries)
        }
        
        logger.info(f"Analysis complete for site {site_id}: {analysis['total_libraries']} libraries, "
                   f"{analysis['total_shared_links']} shared links, {analysis['total_permissions']} permissions")
        
        return analysis
    
    def analyze_all_sites(self) -> List[Dict[str, Any]]:
        """
        Analyze permissions and shared links for all discovered sites
        
        Returns:
            list: List of analysis data for all sites
        """
        if not self.sites:
            logger.info("No sites discovered. Running site discovery...")
            self.discover_all_sites()
        
        logger.info(f"Starting analysis of {len(self.sites)} sites...")
        all_analyses = []
        
        for i, site in enumerate(self.sites, 1):
            site_id = site.get("id")
            site_name = site.get("displayName", "Unknown")
            
            if not site_id:
                logger.warning(f"Skipping site {site_name} - no ID found")
                continue
            
            logger.info(f"Analyzing site {i}/{len(self.sites)}: {site_name}")
            
            try:
                analysis = self.analyze_site_permissions(site_id)
                if analysis:
                    all_analyses.append(analysis)
            except Exception as e:
                logger.error(f"Error analyzing site {site_name}: {str(e)}")
                continue
        
        logger.info(f"Analysis complete for {len(all_analyses)} sites")
        return all_analyses
