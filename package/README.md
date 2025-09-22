# SharePoint MCP Server

A Model Context Protocol server for browsing and interacting with Microsoft SharePoint sites and documents.

This TypeScript-based MCP server provides comprehensive SharePoint integration through Microsoft Graph API, enabling:

- **Resources**: Access SharePoint sites as MCP resources with structured metadata
- **Tools**: Search files, list sites, browse document libraries, and retrieve file content with comprehensive document parsing
- **Document Support**: PDF, Word, Excel, PowerPoint, text files, HTML, RTF, and more
- **Authentication**: Secure OAuth2 client credentials flow with automatic token management

## Features

### Enhanced Document Parsing

The server now supports comprehensive document parsing for multiple file formats:

- **PDF Documents** (.pdf) - Full text extraction with page count and metadata
- **Microsoft Word** (.doc, .docx) - Text content extraction with formatting preservation
- **Microsoft Excel** (.xls, .xlsx) - All worksheets with data in CSV format
- **Microsoft PowerPoint** (.ppt, .pptx) - Slide text content extraction
- **Text Files** (.txt, .md, .json, .xml, .csv) - Direct content reading
- **HTML Files** (.html, .htm) - Clean text extraction without markup
- **RTF Documents** (.rtf) - Basic text extraction
- **Automatic Fallback** - Attempts text parsing for any text-based format

### Resources

- List SharePoint sites accessible to the application
- Access individual site information via `sharepoint://site/{siteId}` URIs
- JSON-formatted site metadata with display names and web URLs

### Tools

#### `search_files`

Search for files and documents within the configured SharePoint site using drive-specific search. This method works with Sites.Selected permissions and is automatically scoped to the configured site and folder.

- **Parameters**:
  - `query` (required): Search query string
  - `limit` (optional): Maximum results to return (default: 10)

#### `list_sites`

List SharePoint sites accessible to the application

- **Parameters**:
  - `search` (optional): Filter sites by display name

#### `get_site_info`

Get detailed information about a specific SharePoint site

- **Parameters**:
  - `siteUrl` (required): SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)

#### `list_site_drives`

List document libraries (drives) in a SharePoint site

- **Parameters**:
  - `siteUrl` (required): SharePoint site URL

#### `list_drive_items`

List files and folders in a SharePoint document library. Uses DEFAULT_SITE_URL and DEFAULT_FOLDER_PATH if configured to reduce noise.

- **Parameters**:
  - `siteUrl` (optional): SharePoint site URL (uses DEFAULT_SITE_URL if not provided)
  - `driveId` (optional): Specific drive ID (uses default drive if not specified)
  - `folderPath` (optional): Folder path to list items from (uses DEFAULT_FOLDER_PATH if available, otherwise root)

#### `get_file_content`

Get the content of a specific file from SharePoint. Supports comprehensive document parsing including PDF, Word (doc/docx), Excel (xls/xlsx), PowerPoint (ppt/pptx), text files, HTML, RTF, and more. Uses DEFAULT_SITE_URL if configured.

- **Parameters**:
  - `filePath` (required): Path to the file
  - `siteUrl` (optional): SharePoint site URL (uses DEFAULT_SITE_URL if not provided)
  - `driveId` (optional): Specific drive ID (uses default drive if not specified)
  - `includeMetadata` (optional): Whether to include document metadata in the response (default: true)

## Installation

### Using npx (Recommended)

You can run the SharePoint MCP Server directly using npx without installing it globally:

```bash
npx @alliottech/sharepoint-mcp-server
```

### Global Installation

Alternatively, you can install it globally:

```bash
npm install -g @alliottech/sharepoint-mcp-server
sharepoint-mcp-server
```

## Prerequisites

### Azure App Registration

1. Register an application in Azure Active Directory
2. Configure API permissions:
   - Microsoft Graph: `Sites.Selected` (Application permission)
3. Grant admin consent for the permissions
4. Create a client secret
5. **Verify site-specific permissions** (required for Sites.Selected):

   The SharePoint administrator must have already granted access to the specific SharePoint site for your app registration. The app will verify access on startup and provide clear error messages if permissions are missing.

### Installation in librechat:

Open `librechat.yaml`

add the following at the bottom of the file:

```bash
mcpServers:
  NAME_OF_MCP:
    command: npx
    args:
      - -y
      - "@alexandersen01/sharepoint-mcp-server-better"
    env:
      SEARCH_REGION: "EMEA"
      DEFAULT_FOLDER_PATH: "${DEFAULT_FOLDER_PATH}"
      DEFAULT_SITE_URL: "${DEFAULT_SITE_URL}"

      TENANT_ID: "${TENANT_ID}"
      CLIENT_ID: "${CLIENT_ID}"
      CLIENT_SECRET: "${CLIENT_SECRET}"
    chatMenu: true
    serverInstructions: |
      SharePoint MCP Server provides access to your SharePoint sites and documents.
      Use this to search, read, and interact with SharePoint content including:
      - Site collections and subsites
      - Document libraries
      - Lists and list items
      - File operations (read, search, metadata)

```

**Remember to run the following command to restart the container**

```bash
ssh -p 22 USER@SSH_ADDR 'cd /home/USER/LibreChat && docker compose -f ./deploy-compose.yml -f ./deploy-compose.override.yml down && docker compose -f ./deploy-compose.yml -f ./deploy-compose.override.yml pull && docker compose -f ./deploy-compose.yml -f ./deploy-compose.override.yml up -d --remove-orphans'
```

If `deploy-compose.override.yml` is not found:

```bash
ssh -p 22 USERs@SSH_ADDR 'cd /home/USER/LibreChat && docker compose -f ./deploy-compose.yml down && docker compose -f ./deploy-compose.yml pull && docker compose -f ./deploy-compose.yml up -d --remove-orphans'
```

### Environment Variables

Set the following environment variables:

```bash
SHAREPOINT_URL=https://yourtenant.sharepoint.com
TENANT_ID=your-azure-tenant-id
CLIENT_ID=your-azure-app-client-id
CLIENT_SECRET=your-azure-app-client-secret

# Optional: Set defaults to reduce noise and focus operations
DEFAULT_SITE_URL=https://yourtenant.sharepoint.com/sites/yoursite
DEFAULT_FOLDER_PATH=Documents/YourFolder
```

#### Site and Folder Filtering

To reduce noise and focus the agent on specific SharePoint sites and folders, you can set:

- **DEFAULT_SITE_URL**: Default SharePoint site URL for operations
- **DEFAULT_FOLDER_PATH**: Default folder path within the site

When these are set, tools like `search_files`, `list_drive_items`, and `get_file_content` will use these defaults when site URL or folder path parameters are not explicitly provided. This helps keep the agent focused on relevant content without overwhelming it with organization-wide SharePoint data.

## Development

Install dependencies:

```bash
npm install
```

Build the server:

```bash
npm run build
```

For development with auto-rebuild:

```bash
npm run watch
```

## Testing

Test the server using the MCP Inspector:

```bash
npm run inspector
```

The Inspector provides a web interface to test all available tools and resources.

## Installation

### Claude Desktop Configuration

Add the server to your Claude Desktop configuration:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "sharepoint-mcp-server": {
      "command": "node",
      "args": ["/path/to/sharepoint-mcp-server/build/index.js"],
      "env": {
        "SHAREPOINT_URL": "https://yourtenant.sharepoint.com",
        "TENANT_ID": "your-azure-tenant-id",
        "CLIENT_ID": "your-azure-app-client-id",
        "CLIENT_SECRET": "your-azure-app-client-secret",
        "DEFAULT_SITE_URL": "https://yourtenant.sharepoint.com/sites/yoursite",
        "DEFAULT_FOLDER_PATH": "Documents/YourFolder"
      }
    }
  }
}
```

### Global Installation

You can also install the server globally:

```bash
npm install -g .
```

Then use it directly:

```bash
sharepoint-mcp-server
```

## Architecture

The server implements a service-oriented architecture with clear separation of concerns:

- **Authentication Layer**: Handles OAuth2 token acquisition and refresh
- **Graph API Client**: Manages HTTP requests to Microsoft Graph API
- **Tool Handlers**: Process MCP tool requests and format responses
- **Resource Handlers**: Manage SharePoint site resources and metadata
- **Error Handling**: Comprehensive error management with proper MCP error codes

## Security Considerations

- Uses OAuth2 client credentials flow for secure authentication
- Tokens are automatically refreshed before expiration
- All API requests use HTTPS
- Client secrets should be stored securely and never committed to version control
- Application permissions require admin consent in Azure AD

### Sites.Selected vs Sites.ReadAll/Files.ReadAll

This server is configured to work with `Sites.Selected` permission for enhanced security:

**Sites.Selected Benefits:**

- ✅ Follows principle of least privilege
- ✅ Only grants access to explicitly allowed SharePoint sites
- ✅ Reduces security risk and compliance concerns
- ✅ Better for enterprise environments

**Migration from Sites.ReadAll/Files.ReadAll:**
If you're upgrading from broader permissions:

1. **Update Azure App Registration:**

   - Remove `Sites.Read.All` and `Files.Read.All` permissions
   - Add `Sites.Selected` permission
   - Grant admin consent

2. **Verify Site-Specific Access:**

   - Ensure your SharePoint administrator has granted the app access to the required sites
   - Only sites with explicit permissions will be accessible

3. **Expected Changes:**
   - App will only access the configured `DEFAULT_SITE_URL`
   - Access to other SharePoint sites will be denied (403 errors)
   - More detailed error messages guide you through permission setup

**Important:** The `DEFAULT_SITE_URL` must be set and the app must have explicit permissions to that site, or the server will fail to start with clear guidance on how to fix it.

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Verify Azure app registration and permissions
2. **Site Access**: Ensure the app has appropriate SharePoint permissions
3. **Network Issues**: Check firewall settings for Microsoft Graph API access
4. **Sites.Selected Permission Issues**:
   - **403 Forbidden errors**: App doesn't have access to the specific SharePoint site
   - **"Security initialization failed"**: Site permissions not granted yet
   - **Solution**: Contact your SharePoint administrator to verify site permissions have been granted
5. **Server fails to start**: Check that `DEFAULT_SITE_URL` is set and accessible

### Debug Mode

Set environment variable for detailed logging:

```bash
DEBUG=sharepoint-mcp-server
```

## Contributing

1. Follow TypeScript best practices
2. Maintain comprehensive error handling
3. Add tests for new functionality
4. Update documentation for API changes

## License

This project is licensed under the Mozilla Public License 2.0. See the [LICENSE](LICENSE) file for details.

## Contributing

We welcome contributions! Please follow these guidelines:

### Getting Started

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass (`npm test`)
6. Commit your changes (`git commit -m 'Add amazing feature'`)
7. Push to the branch (`git push origin feature/amazing-feature`)
8. Open a Pull Request

### Development Guidelines

- Follow TypeScript best practices and maintain type safety
- Implement comprehensive error handling with proper MCP error codes
- Add JSDoc comments for all public methods and classes
- Maintain the service-oriented architecture with clear separation of concerns
- Follow SOLID principles and keep functions focused and testable
- Update documentation for any API changes

### Code Style

- Use TypeScript strict mode
- Follow the existing code formatting and naming conventions
- Remove unused imports and variables
- Use descriptive variable and function names
- Prefer composition over inheritance

### Testing

- Add unit tests for new functionality
- Test error conditions and edge cases
- Ensure the basic test suite passes
- Test with real SharePoint environments when possible

### Documentation

- Update README.md for new features or configuration changes
- Add JSDoc comments for new public APIs
- Include examples for complex functionality
- Update the changelog for significant changes

## Changelog

### [0.2.1] - Site and Folder Filtering

- **NEW**: DEFAULT_SITE_URL environment variable for scoping operations to specific SharePoint site
- **NEW**: DEFAULT_FOLDER_PATH environment variable for scoping operations to specific folder
- **ENHANCED**: `search_files` tool now supports site and folder filtering to reduce noise
- **ENHANCED**: `list_drive_items` tool now uses default site and folder when not specified
- **ENHANCED**: `get_file_content` tool now uses default site when not specified
- **IMPROVED**: Tool parameter requirements relaxed when defaults are configured
- **IMPROVED**: Search results filtered to ensure they match the specified site

### [0.2.0] - Enhanced Document Parsing

- **NEW**: Comprehensive document parsing support for multiple file formats
- **NEW**: PDF document text extraction with metadata
- **NEW**: Microsoft Word (.doc, .docx) content parsing
- **NEW**: Microsoft Excel (.xls, .xlsx) data extraction from all worksheets
- **NEW**: Microsoft PowerPoint (.ppt, .pptx) slide text extraction
- **NEW**: HTML document parsing with clean text extraction
- **NEW**: RTF document basic text extraction
- **ENHANCED**: `get_file_content` tool now supports multiple document formats
- **ENHANCED**: Added `includeMetadata` parameter for detailed document information
- **ADDED**: Dependencies: pdf-parse, mammoth, xlsx, yauzl, node-html-parser

### [0.1.0] - Initial Release

- Basic SharePoint integration via Microsoft Graph API
- Support for searching files across SharePoint
- Site listing and browsing capabilities
- Document library access and file content retrieval (text files only)
- OAuth2 client credentials authentication
- MCP resource support for SharePoint sites
- Comprehensive error handling and logging

## Support

If you encounter issues or have questions:

1. Check the [troubleshooting section](#troubleshooting) in this README
2. Search existing [GitHub issues](https://github.com/alexandersen01/sharepoint-mcp-server-better/issues)
3. Create a new issue with detailed information about your problem
4. Include relevant logs and configuration (without sensitive information)

## Acknowledgments

- Built with the [Model Context Protocol SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- Uses Microsoft Graph API for SharePoint integration
- Inspired by the MCP community and ecosystem
