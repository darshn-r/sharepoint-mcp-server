#!/usr/bin/env python3
"""
SharePoint MCP Server
A Model Context Protocol server for interacting with SharePoint Online
"""

import asyncio
import os
import requests
from typing import Any, Dict, List, Optional
from fastmcp import FastMCP
from pydantic import BaseModel

class SharePointConfig(BaseModel):
    site_id: str
    client_id: str
    client_secret: str
    tenant_id: str

class SharePointMCPServer:
    def __init__(self):
        self.config = SharePointConfig(
            site_id=os.getenv("SHAREPOINT_SITE_ID", ""),
            client_id=os.getenv("SHAREPOINT_CLIENT_ID", ""),
            client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET", ""),
            tenant_id=os.getenv("SHAREPOINT_TENANT_ID", "")
        )
        self.access_token = None
        
    def get_access_token(self) -> Optional[str]:
        """Get OAuth2 access token for SharePoint"""
        try:
            resource = "https://graph.microsoft.com/"
            grant_type = "client_credentials"
            token_api = f"https://login.microsoftonline.com/{self.config.tenant_id}/oauth2/token"
            
            payload = f"grant_type={grant_type}&client_id={self.config.client_id}&client_secret={self.config.client_secret}&resource={resource}"
            
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            response = requests.post(token_api, data=payload, headers=headers, verify=True)
            response.raise_for_status()
            
            token_data = response.json()
            self.access_token = token_data['access_token']
            return self.access_token
            
        except Exception as e:
            print(f"Error getting access token: {e}")
            return None
    
    def get_site_files(self, max_files: int = 100, folder_id: Optional[str] = None) -> List[Dict[str, Any]]:
        """Get files from the SharePoint site with pagination and limits"""
        if not self.access_token:
            if not self.get_access_token():
                return []
        
        try:
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            # Build URL based on whether we're looking at root or specific folder
            if folder_id:
                files_url = f"https://graph.microsoft.com/v1.0/sites/{self.config.site_id}/drive/items/{folder_id}/children"
            else:
                files_url = f"https://graph.microsoft.com/v1.0/sites/{self.config.site_id}/drive/root/children"
            
            # Add pagination and ordering
            files_url += f"?$top={min(max_files, 100)}&$orderby=lastModifiedDateTime desc"
            
            files_response = requests.get(files_url, headers=headers)
            files_response.raise_for_status()
            
            files_data = files_response.json()
            files = []
            
            # Process files and folders (but don't recurse automatically)
            for item in files_data.get('value', []):
                file_info = {
                    'name': item.get('name', ''),
                    'type': 'folder' if 'folder' in item else 'file',
                    'size': item.get('size', 0),
                    'created': item.get('createdDateTime', ''),
                    'modified': item.get('lastModifiedDateTime', ''),
                    'id': item.get('id', ''),
                    'webUrl': item.get('webUrl', ''),
                    'parentPath': item.get('parentReference', {}).get('path', ''),
                    'mimeType': item.get('file', {}).get('mimeType', '') if 'file' in item else ''
                }
                files.append(file_info)
                
                # Stop if we've reached the max files limit
                if len(files) >= max_files:
                    break
            
            return files
            
        except Exception as e:
            print(f"Error getting site files: {e}")
            return [{'error': f'Failed to retrieve files: {str(e)}'}]
    
    def get_folder_contents(self, folder_id: str, max_files: int = 50) -> List[Dict[str, Any]]:
        """Get contents of a specific folder"""
        if not self.access_token:
            if not self.get_access_token():
                return []
        
        try:
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            folder_url = f"https://graph.microsoft.com/v1.0/sites/{self.config.site_id}/drive/items/{folder_id}/children"
            folder_url += f"?$top={min(max_files, 100)}&$orderby=lastModifiedDateTime desc"
            
            response = requests.get(folder_url, headers=headers)
            response.raise_for_status()
            
            folder_data = response.json()
            files = []
            
            for item in folder_data.get('value', []):
                file_info = {
                    'name': item.get('name', ''),
                    'type': 'folder' if 'folder' in item else 'file',
                    'size': item.get('size', 0),
                    'created': item.get('createdDateTime', ''),
                    'modified': item.get('lastModifiedDateTime', ''),
                    'id': item.get('id', ''),
                    'webUrl': item.get('webUrl', ''),
                    'parentFolder': folder_id,
                    'parentPath': item.get('parentReference', {}).get('path', ''),
                    'mimeType': item.get('file', {}).get('mimeType', '') if 'file' in item else ''
                }
                files.append(file_info)
                
                if len(files) >= max_files:
                    break
                    
            return files
            
        except Exception as e:
            print(f"Error getting folder contents: {e}")
            return [{'error': f'Failed to retrieve folder contents: {str(e)}'}]
    
    def get_file_content(self, file_id: str, max_size: int = 5000) -> Optional[str]:
        """Get content of a specific file with size limits"""
        if not self.access_token:
            if not self.get_access_token():
                return None
        
        try:
            # First get file metadata to check size
            file_info_url = f"https://graph.microsoft.com/v1.0/sites/{self.config.site_id}/drive/items/{file_id}"
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            info_response = requests.get(file_info_url, headers=headers)
            info_response.raise_for_status()
            
            file_info = info_response.json()
            file_size = file_info.get('size', 0)
            file_name = file_info.get('name', 'unknown')
            mime_type = file_info.get('file', {}).get('mimeType', '')
            
            # Check if file is too large
            if file_size > max_size * 1024:  # max_size in KB
                return f"File '{file_name}' is too large ({file_size} bytes). Maximum allowed size is {max_size}KB."
            
            # Check if file type is readable
            text_types = ['text/', 'application/json', 'application/xml', 'application/javascript', 'application/csv']
            if mime_type and not any(text_type in mime_type for text_type in text_types):
                return f"File '{file_name}' is a binary file ({mime_type}). Only text files can be read."
            
            # Get file content
            file_url = f"https://graph.microsoft.com/v1.0/sites/{self.config.site_id}/drive/items/{file_id}/content"
            headers = {
                'Authorization': f'Bearer {self.access_token}'
            }
            
            response = requests.get(file_url, headers=headers)
            response.raise_for_status()
            
            # Try to decode as text
            try:
                content = response.text
                # Truncate if still too long
                if len(content) > max_size:
                    content = content[:max_size] + f"\n\n... (truncated, showing first {max_size} characters)"
                return content
            except UnicodeDecodeError:
                return f"File '{file_name}' contains binary data that cannot be displayed as text."
                
        except Exception as e:
            print(f"Error getting file content: {e}")
            return f"Error retrieving file content: {str(e)}"

# Initialize the SharePoint server
sharepoint_server = SharePointMCPServer()

# Create FastMCP instance
mcp = FastMCP("SharePoint MCP Server")

@mcp.tool()
def list_sharepoint_files(max_files: int = 50, folder_id: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    List files in the SharePoint site with pagination.
    
    Args:
        max_files: Maximum number of files to return (default: 50, max: 100)
        folder_id: Optional folder ID to list contents of specific folder
    
    Returns:
        List of files with their metadata including name, type, size, dates, and URLs.
    """
    # Limit max_files to prevent overwhelming responses
    max_files = min(max_files, 100)
    files = sharepoint_server.get_site_files(max_files=max_files, folder_id=folder_id)
    return files

@mcp.tool()
def get_folder_contents(folder_id: str, max_files: int = 50) -> List[Dict[str, Any]]:
    """
    Get contents of a specific SharePoint folder.
    
    Args:
        folder_id: The ID of the folder to list contents for
        max_files: Maximum number of files to return (default: 50, max: 100)
    
    Returns:
        List of files and subfolders in the specified folder.
    """
    max_files = min(max_files, 100)
    return sharepoint_server.get_folder_contents(folder_id, max_files)

@mcp.tool()
def get_sharepoint_file_content(file_id: str, max_size_kb: int = 5) -> Optional[str]:
    """
    Get the content of a specific SharePoint file.
    
    Args:
        file_id: The ID of the file to retrieve content for
        max_size_kb: Maximum file size in KB to read (default: 5KB, max: 50KB)
        
    Returns:
        The file content as text (with size limits for safety)
    """
    # Limit max size to prevent overwhelming responses
    max_size_kb = min(max_size_kb, 50)
    content = sharepoint_server.get_file_content(file_id, max_size_kb * 1024)
    return content

@mcp.tool()
def test_sharepoint_connection() -> Dict[str, Any]:
    """
    Test the connection to SharePoint by attempting to get an access token.
    
    Returns:
        Connection status and configuration info
    """
    token = sharepoint_server.get_access_token()
    return {
        "connected": token is not None,
        "site_id": sharepoint_server.config.site_id,
        "tenant_id": sharepoint_server.config.tenant_id,
        "client_id": sharepoint_server.config.client_id,
        "token_available": sharepoint_server.access_token is not None
    }

if __name__ == "__main__":
    # Check if environment variables are set
    required_vars = ["SHAREPOINT_SITE_ID", "SHAREPOINT_CLIENT_ID", "SHAREPOINT_CLIENT_SECRET", "SHAREPOINT_TENANT_ID"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        print(f"Missing required environment variables: {', '.join(missing_vars)}")
        print("Please set these environment variables before running the server.")
        exit(1)
    
    # Run the server
    mcp.run()