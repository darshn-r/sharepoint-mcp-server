# SharePoint MCP Server

A Model Context Protocol (MCP) server for interacting with SharePoint Online using Microsoft Graph API. This server provides secure access to SharePoint files and folders with built-in safety limits and authentication.

## Features

- **File Management**: List files and folders in SharePoint sites
- **Content Access**: Read text-based file contents with size limits
- **Folder Navigation**: Browse folder structures and contents
- **Safety Limits**: Built-in protections against large file downloads
- **Authentication**: OAuth2 client credentials flow with Microsoft Graph
- **Pagination**: Efficient handling of large file lists

## Prerequisites

- Python 3.7+
- SharePoint Online site
- Azure App Registration with appropriate permissions
- Required Python packages (see Installation)

## Installation

1. Clone or download the SharePoint MCP Server code
2. Install required dependencies:

```bash
pip install fastmcp requests pydantic
```

## Azure App Registration Setup

Before using the server, you need to register an application in Azure AD:

1. **Go to Azure Portal** → Azure Active Directory → App registrations
2. **Create new registration**:
   - Name: "SharePoint MCP Server" (or your preferred name)
   - Supported account types: Single tenant
   - Redirect URI: Not required for this setup

3. **Note the Application (client) ID** and **Directory (tenant) ID**

4. **Create client secret**:
   - Go to "Certificates & secrets"
   - Create new client secret
   - Copy the secret value (you won't see it again)

5. **Configure API permissions**:
   - Go to "API permissions"
   - Add permissions → Microsoft Graph → Application permissions
   - Add these permissions:
     - `Sites.Read.All` (to read SharePoint sites)
     - `Files.Read.All` (to read files)
   - **Grant admin consent** for your organization

## Configuration

Set the following environment variables:

```bash
export SHAREPOINT_SITE_ID="your-site-id"
export SHAREPOINT_CLIENT_ID="your-client-id"
export SHAREPOINT_CLIENT_SECRET="your-client-secret"
export SHAREPOINT_TENANT_ID="your-tenant-id"
```

### Finding Your SharePoint Site ID

You can find your SharePoint site ID using Microsoft Graph Explorer or by using the following URL format:
```
https://graph.microsoft.com/v1.0/sites/{hostname}:/{site-path}
```

For example:
```
https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/marketing
```

## Usage

### Running the Server

```bash
python sharepoint_mcp_server.py
```

The server will start and be available for MCP client connections.

### Available Tools

#### 1. `list_sharepoint_files`
Lists files and folders in the SharePoint site.

**Parameters:**
- `max_files` (int, optional): Maximum number of files to return (default: 50, max: 100)
- `folder_id` (str, optional): ID of specific folder to list contents

**Returns:** List of files with metadata including name, type, size, dates, and URLs.

#### 2. `get_folder_contents`
Gets contents of a specific SharePoint folder.

**Parameters:**
- `folder_id` (str, required): The ID of the folder to list contents for
- `max_files` (int, optional): Maximum number of files to return (default: 50, max: 100)

**Returns:** List of files and subfolders in the specified folder.

#### 3. `get_sharepoint_file_content`
Retrieves the content of a specific SharePoint file.

**Parameters:**
- `file_id` (str, required): The ID of the file to retrieve content for
- `max_size_kb` (int, optional): Maximum file size in KB to read (default: 5KB, max: 50KB)

**Returns:** The file content as text (with size limits for safety).

**Note:** Only text-based files are supported. Binary files will return an appropriate message.

#### 4. `test_sharepoint_connection`
Tests the connection to SharePoint by attempting to get an access token.

**Parameters:** None

**Returns:** Connection status and configuration info.

## Security Features

- **Size Limits**: File content reading is limited to prevent memory issues
- **File Type Filtering**: Only text-based files can be read
- **Authentication**: Uses OAuth2 client credentials flow
- **Pagination**: Limits the number of files returned per request
- **Error Handling**: Graceful handling of authentication and API errors

## Supported File Types

The server can read content from these file types:
- Plain text files (`.txt`, `.md`, `.csv`, etc.)
- JSON files (`.json`)
- XML files (`.xml`)
- JavaScript files (`.js`)
- Configuration files
- Any file with MIME type starting with `text/`

## Error Handling

The server includes comprehensive error handling for:
- Missing environment variables
- Authentication failures
- File access errors
- Network connectivity issues
- File size and type restrictions

## Troubleshooting

### Common Issues

1. **Authentication Error**: 
   - Verify all environment variables are set correctly
   - Check that the Azure app has the required permissions
   - Ensure admin consent has been granted

2. **Site Not Found**:
   - Verify the SharePoint site ID is correct
   - Check that the app has access to the site

3. **File Access Denied**:
   - Ensure the app has `Files.Read.All` permission
   - Check that the file exists and is accessible

4. **Connection Test Fails**:
   - Run `test_sharepoint_connection()` to diagnose issues
   - Check network connectivity to Microsoft Graph

### Debug Mode

To enable debug logging, you can modify the error handling in the code to include more detailed error messages.

## Limitations

- **Read-Only**: This server only supports reading files, not writing or modifying
- **Text Files Only**: Binary files cannot be read (by design for security)
- **Size Limits**: File content is limited to prevent memory issues
- **Application Permissions**: Uses app-only authentication (no user context)

## License

This project is provided as-is for educational and development purposes. Please ensure compliance with your organization's policies and Microsoft's terms of service.

## Contributing

To contribute to this project:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Support

For issues related to:
- **Microsoft Graph API**: Check the [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph/)
- **SharePoint**: Refer to [SharePoint documentation](https://docs.microsoft.com/en-us/sharepoint/)
- **MCP Protocol**: See the [Model Context Protocol specification](https://github.com/modelcontextprotocol/python-sdk)

---

**Note**: This server requires proper Azure AD application setup and SharePoint permissions. Always follow your organization's security policies when configuring access to SharePoint resources.