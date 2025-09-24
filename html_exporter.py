"""
HTML Export Module
Handles exporting SharePoint analysis data to formatted HTML reports
"""

import json
from datetime import datetime
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)

class HTMLExporter:
    """Handles exporting SharePoint analysis data to HTML format"""
    
    def __init__(self):
        self.template = self._get_html_template()
    
    def _get_html_template(self) -> str:
        """Get the HTML template with CSS styling"""
        return """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SharePoint Permissions Analysis Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }
        
        .header p {
            margin: 10px 0 0 0;
            opacity: 0.9;
            font-size: 1.1em;
        }
        
        .summary {
            padding: 30px;
            background-color: #f8f9fa;
            border-bottom: 1px solid #e9ecef;
        }
        
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        
        .summary-card {
            background: white;
            padding: 20px;
            border-radius: 6px;
            text-align: center;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        
        .summary-card h3 {
            margin: 0 0 10px 0;
            color: #0078d4;
            font-size: 2em;
        }
        
        .summary-card p {
            margin: 0;
            color: #666;
            font-size: 0.9em;
        }
        
        .content {
            padding: 30px;
        }
        
        .site-section {
            margin-bottom: 40px;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            overflow: hidden;
        }
        
        .site-header {
            background-color: #0078d4;
            color: white;
            padding: 20px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .site-header:hover {
            background-color: #106ebe;
        }
        
        .site-header h2 {
            margin: 0;
            font-size: 1.5em;
        }
        
        .site-header .toggle {
            font-size: 1.2em;
            transition: transform 0.3s;
        }
        
        .site-header.collapsed .toggle {
            transform: rotate(-90deg);
        }
        
        .site-content {
            padding: 20px;
            display: none;
        }
        
        .site-content.expanded {
            display: block;
        }
        
        .library-section {
            margin-bottom: 30px;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            overflow: hidden;
        }
        
        .library-header {
            background-color: #f8f9fa;
            padding: 15px;
            border-bottom: 1px solid #e9ecef;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .library-header:hover {
            background-color: #e9ecef;
        }
        
        .library-header h3 {
            margin: 0;
            color: #0078d4;
            font-size: 1.2em;
        }
        
        .library-content {
            padding: 20px;
            display: none;
        }
        
        .library-content.expanded {
            display: block;
        }
        
        .section-title {
            color: #0078d4;
            font-size: 1.1em;
            margin: 20px 0 10px 0;
            padding-bottom: 5px;
            border-bottom: 2px solid #0078d4;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
            background-color: white;
        }
        
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #e9ecef;
        }
        
        th {
            background-color: #f8f9fa;
            font-weight: 600;
            color: #495057;
        }
        
        tr:hover {
            background-color: #f8f9fa;
        }
        
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: 500;
            text-transform: uppercase;
        }
        
        .badge-read {
            background-color: #d4edda;
            color: #155724;
        }
        
        .badge-write {
            background-color: #fff3cd;
            color: #856404;
        }
        
        .badge-admin {
            background-color: #f8d7da;
            color: #721c24;
        }
        
        .badge-owner {
            background-color: #cce5ff;
            color: #004085;
        }
        
        .link {
            color: #0078d4;
            text-decoration: none;
        }
        
        .link:hover {
            text-decoration: underline;
        }
        
        .no-data {
            text-align: center;
            color: #6c757d;
            font-style: italic;
            padding: 20px;
        }
        
        .footer {
            background-color: #f8f9fa;
            padding: 20px;
            text-align: center;
            color: #6c757d;
            border-top: 1px solid #e9ecef;
        }
        
        .loading {
            text-align: center;
            padding: 40px;
            color: #6c757d;
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 0;
            }
            
            .header {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .content {
                padding: 20px;
            }
            
            .summary-grid {
                grid-template-columns: 1fr;
            }
            
            table {
                font-size: 0.9em;
            }
            
            th, td {
                padding: 8px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>SharePoint Permissions Analysis</h1>
            <p>Comprehensive report of sites, libraries, shared links, and permissions</p>
        </div>
        
        <div class="summary">
            <h2>Executive Summary</h2>
            <div class="summary-grid">
                <div class="summary-card">
                    <h3 id="total-sites">0</h3>
                    <p>Total Sites</p>
                </div>
                <div class="summary-card">
                    <h3 id="total-libraries">0</h3>
                    <p>Total Libraries</p>
                </div>
                <div class="summary-card">
                    <h3 id="total-shared-links">0</h3>
                    <p>Shared Links</p>
                </div>
                <div class="summary-card">
                    <h3 id="total-permissions">0</h3>
                    <p>Total Permissions</p>
                </div>
            </div>
        </div>
        
        <div class="content">
            <div id="sites-container">
                <!-- Sites will be populated here -->
            </div>
        </div>
        
        <div class="footer">
            <p>Report generated on {timestamp} | SharePoint Permissions Analyzer</p>
        </div>
    </div>
    
    <script>
        // Toggle functionality for collapsible sections
        document.addEventListener('DOMContentLoaded', function() {
            // Site toggle functionality
            document.addEventListener('click', function(e) {
                if (e.target.closest('.site-header')) {
                    const siteHeader = e.target.closest('.site-header');
                    const siteContent = siteHeader.nextElementSibling;
                    const toggle = siteHeader.querySelector('.toggle');
                    
                    if (siteContent.classList.contains('expanded')) {
                        siteContent.classList.remove('expanded');
                        siteHeader.classList.add('collapsed');
                    } else {
                        siteContent.classList.add('expanded');
                        siteHeader.classList.remove('collapsed');
                    }
                }
                
                // Library toggle functionality
                if (e.target.closest('.library-header')) {
                    const libraryHeader = e.target.closest('.library-header');
                    const libraryContent = libraryHeader.nextElementSibling;
                    
                    if (libraryContent.classList.contains('expanded')) {
                        libraryContent.classList.remove('expanded');
                    } else {
                        libraryContent.classList.add('expanded');
                    }
                }
            });
        });
        
        function getRoleBadgeClass(roles) {
            if (!roles || roles.length === 0) return 'badge-read';
            
            const roleStr = roles.join(',').toLowerCase();
            if (roleStr.includes('owner') || roleStr.includes('full')) {
                return 'badge-owner';
            } else if (roleStr.includes('admin') || roleStr.includes('manage')) {
                return 'badge-admin';
            } else if (roleStr.includes('write') || roleStr.includes('edit')) {
                return 'badge-write';
            } else {
                return 'badge-read';
            }
        }
    </script>
</body>
</html>
        """
    
    def export_to_html(self, analysis_data: List[Dict[str, Any]], output_file: str = "sharepoint_analysis_report.html") -> str:
        """
        Export analysis data to HTML file
        
        Args:
            analysis_data: List of site analysis data
            output_file: Output HTML file path
            
        Returns:
            str: Path to the generated HTML file
        """
        logger.info(f"Exporting analysis data to HTML: {output_file}")
        
        # Calculate summary statistics
        total_sites = len(analysis_data)
        total_libraries = sum(site.get("total_libraries", 0) for site in analysis_data)
        total_shared_links = sum(site.get("total_shared_links", 0) for site in analysis_data)
        total_permissions = sum(site.get("total_permissions", 0) for site in analysis_data)
        
        # Generate HTML content
        html_content = self.template.replace("{timestamp}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        # Replace summary statistics
        html_content = html_content.replace('id="total-sites">0', f'id="total-sites">{total_sites}')
        html_content = html_content.replace('id="total-libraries">0', f'id="total-libraries">{total_libraries}')
        html_content = html_content.replace('id="total-shared-links">0', f'id="total-shared-links">{total_shared_links}')
        html_content = html_content.replace('id="total-permissions">0', f'id="total-permissions">{total_permissions}')
        
        # Generate sites HTML
        sites_html = self._generate_sites_html(analysis_data)
        html_content = html_content.replace('<!-- Sites will be populated here -->', sites_html)
        
        # Write to file
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logger.info(f"HTML report successfully exported to: {output_file}")
            return output_file
            
        except Exception as e:
            logger.error(f"Error writing HTML file: {str(e)}")
            raise
    
    def _generate_sites_html(self, analysis_data: List[Dict[str, Any]]) -> str:
        """Generate HTML for all sites"""
        sites_html = ""
        
        for site_data in analysis_data:
            site_info = site_data.get("site_info", {})
            libraries = site_data.get("libraries", [])
            
            site_name = site_info.get("displayName", "Unknown Site")
            site_url = site_info.get("webUrl", "#")
            site_id = site_info.get("id", "")
            
            sites_html += f"""
            <div class="site-section">
                <div class="site-header collapsed">
                    <h2>{site_name}</h2>
                    <span class="toggle">▼</span>
                </div>
                <div class="site-content">
                    <p><strong>Site URL:</strong> <a href="{site_url}" target="_blank" class="link">{site_url}</a></p>
                    <p><strong>Site ID:</strong> {site_id}</p>
                    <p><strong>Created:</strong> {site_info.get('createdDateTime', 'Unknown')}</p>
                    <p><strong>Last Modified:</strong> {site_info.get('lastModifiedDateTime', 'Unknown')}</p>
                    
                    <div class="section-title">Libraries ({len(libraries)})</div>
                    {self._generate_libraries_html(libraries)}
                </div>
            </div>
            """
        
        return sites_html
    
    def _generate_libraries_html(self, libraries: List[Dict[str, Any]]) -> str:
        """Generate HTML for libraries within a site"""
        if not libraries:
            return '<div class="no-data">No libraries found</div>'
        
        libraries_html = ""
        
        for library in libraries:
            library_name = library.get("name", "Unknown Library")
            library_url = library.get("webUrl", "#")
            shared_links = library.get("shared_links", [])
            permissions = library.get("permissions", [])
            
            libraries_html += f"""
            <div class="library-section">
                <div class="library-header">
                    <h3>{library_name}</h3>
                    <span class="toggle">▼</span>
                </div>
                <div class="library-content">
                    <p><strong>Library URL:</strong> <a href="{library_url}" target="_blank" class="link">{library_url}</a></p>
                    <p><strong>Description:</strong> {library.get('description', 'No description')}</p>
                    <p><strong>Created:</strong> {library.get('createdDateTime', 'Unknown')}</p>
                    <p><strong>Last Modified:</strong> {library.get('lastModifiedDateTime', 'Unknown')}</p>
                    
                    {self._generate_shared_links_html(shared_links)}
                    {self._generate_permissions_html(permissions)}
                </div>
            </div>
            """
        
        return libraries_html
    
    def _generate_shared_links_html(self, shared_links: List[Dict[str, Any]]) -> str:
        """Generate HTML for shared links"""
        if not shared_links:
            return '<div class="section-title">Shared Links</div><div class="no-data">No shared links found</div>'
        
        html = f'<div class="section-title">Shared Links ({len(shared_links)})</div>'
        html += '<table><thead><tr><th>Name</th><th>URL</th><th>Size</th><th>Created By</th><th>Created Date</th><th>Last Modified</th></tr></thead><tbody>'
        
        for link in shared_links:
            name = link.get("name", "Unknown")
            url = link.get("webUrl", "#")
            size = self._format_file_size(link.get("size", 0))
            created_by = link.get("createdBy", {}).get("user", {}).get("displayName", "Unknown")
            created_date = link.get("createdDateTime", "Unknown")
            last_modified = link.get("lastModifiedDateTime", "Unknown")
            
            html += f"""
            <tr>
                <td>{name}</td>
                <td><a href="{url}" target="_blank" class="link">View File</a></td>
                <td>{size}</td>
                <td>{created_by}</td>
                <td>{created_date}</td>
                <td>{last_modified}</td>
            </tr>
            """
        
        html += '</tbody></table>'
        return html
    
    def _generate_permissions_html(self, permissions: List[Dict[str, Any]]) -> str:
        """Generate HTML for permissions"""
        if not permissions:
            return '<div class="section-title">Permissions</div><div class="no-data">No permissions found</div>'
        
        html = f'<div class="section-title">Permissions ({len(permissions)})</div>'
        html += '<table><thead><tr><th>Granted To</th><th>Roles</th><th>Type</th><th>Expiration</th><th>Inherited From</th></tr></thead><tbody>'
        
        for permission in permissions:
            roles = permission.get("roles", [])
            granted_to = permission.get("grantedTo", {})
            granted_to_identities = permission.get("grantedToIdentities", [])
            link = permission.get("link", {})
            expiration = permission.get("expirationDateTime", "Never")
            inherited_from = permission.get("inheritedFrom", {})
            
            # Determine who the permission is granted to
            granted_to_name = "Unknown"
            if granted_to:
                granted_to_name = granted_to.get("user", {}).get("displayName", "Unknown User")
            elif granted_to_identities:
                granted_to_name = granted_to_identities[0].get("user", {}).get("displayName", "Unknown User")
            
            # Determine permission type
            perm_type = "Direct"
            if inherited_from:
                perm_type = "Inherited"
            elif link:
                perm_type = "Link"
            
            # Format roles
            roles_html = ""
            for role in roles:
                badge_class = self._get_role_badge_class(role)
                roles_html += f'<span class="badge {badge_class}">{role}</span> '
            
            html += f"""
            <tr>
                <td>{granted_to_name}</td>
                <td>{roles_html}</td>
                <td>{perm_type}</td>
                <td>{expiration}</td>
                <td>{inherited_from.get('drive', {}).get('name', 'N/A')}</td>
            </tr>
            """
        
        html += '</tbody></table>'
        return html
    
    def _format_file_size(self, size_bytes: int) -> str:
        """Format file size in human readable format"""
        if size_bytes == 0:
            return "0 B"
        
        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def _get_role_badge_class(self, role: str) -> str:
        """Get CSS class for role badge"""
        role_lower = role.lower()
        if "owner" in role_lower or "full" in role_lower:
            return "badge-owner"
        elif "admin" in role_lower or "manage" in role_lower:
            return "badge-admin"
        elif "write" in role_lower or "edit" in role_lower:
            return "badge-write"
        else:
            return "badge-read"
