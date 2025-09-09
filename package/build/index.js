#!/usr/bin/env node
/**
 * SharePoint MCP Server
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema, ListResourcesRequestSchema, ReadResourceRequestSchema, McpError, ErrorCode, } from "@modelcontextprotocol/sdk/types.js";
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import yauzl from 'yauzl';
import { parse as parseHtml } from 'node-html-parser';
/**
 * Environment variables required for SharePoint authentication
 */
const { SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET, DEFAULT_SITE_URL, DEFAULT_FOLDER_PATH } = process.env;
if (!SHAREPOINT_URL || !TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error("Required environment variables: SHAREPOINT_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET");
}
/**
 * Document parser for various file formats
 */
class DocumentParser {
    /**
     * Parse document content based on file extension and MIME type
     */
    static async parseDocument(buffer, filename, mimeType) {
        const extension = filename.toLowerCase().split('.').pop();
        
        try {
            switch (extension) {
                case 'pdf':
                    return await this.parsePDF(buffer);
                case 'doc':
                case 'docx':
                    return await this.parseWord(buffer);
                case 'xls':
                case 'xlsx':
                    return await this.parseExcel(buffer);
                case 'ppt':
                case 'pptx':
                    return await this.parsePowerPoint(buffer);
                case 'txt':
                case 'md':
                case 'json':
                case 'xml':
                case 'csv':
                    return this.parseText(buffer);
                case 'html':
                case 'htm':
                    return this.parseHTML(buffer);
                case 'rtf':
                    return this.parseRTF(buffer);
                default:
                    // Try to parse as text if it's a text-based format
                    if (mimeType && mimeType.startsWith('text/')) {
                        return this.parseText(buffer);
                    }
                    throw new Error(`Unsupported file format: ${extension}`);
            }
        } catch (error) {
            throw new Error(`Failed to parse ${extension} file: ${error.message}`);
        }
    }

    /**
     * Parse PDF documents
     */
    static async parsePDF(buffer) {
        // Dynamic import to avoid debug mode issues
        const { default: pdfParse } = await import('pdf-parse');
        const data = await pdfParse(buffer);
        return {
            text: data.text,
            metadata: {
                pages: data.numpages,
                info: data.info
            }
        };
    }

    /**
     * Parse Word documents (.doc, .docx)
     */
    static async parseWord(buffer) {
        const result = await mammoth.extractRawText({ buffer });
        return {
            text: result.value,
            metadata: {
                messages: result.messages
            }
        };
    }

    /**
     * Parse Excel spreadsheets (.xls, .xlsx)
     */
    static async parseExcel(buffer) {
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        let text = '';
        let metadata = {
            sheets: [],
            totalRows: 0,
            totalCells: 0
        };

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const csvData = XLSX.utils.sheet_to_csv(sheet);
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            
            text += `\n\n=== Sheet: ${sheetName} ===\n`;
            text += csvData;
            
            metadata.sheets.push({
                name: sheetName,
                rows: jsonData.length,
                range: sheet['!ref']
            });
            metadata.totalRows += jsonData.length;
        });

        return {
            text: text.trim(),
            metadata
        };
    }

    /**
     * Parse PowerPoint presentations (.ppt, .pptx)
     */
    static async parsePowerPoint(buffer) {
        return new Promise((resolve, reject) => {
            yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
                if (err) {
                    // If it's not a ZIP file, try to extract text as best we can
                    resolve({
                        text: "[PowerPoint content - binary format not fully supported]",
                        metadata: { note: "PowerPoint parsing requires additional libraries for full support" }
                    });
                    return;
                }

                let slides = [];
                let slideText = '';

                zipfile.on("entry", (entry) => {
                    if (entry.fileName.includes("slide") && entry.fileName.endsWith(".xml")) {
                        zipfile.openReadStream(entry, (err, readStream) => {
                            if (err) {
                                zipfile.readEntry();
                                return;
                            }

                            let data = '';
                            readStream.on('data', (chunk) => {
                                data += chunk;
                            });

                            readStream.on('end', () => {
                                // Extract text from XML
                                const textMatches = data.match(/<a:t[^>]*>([^<]*)<\/a:t>/g);
                                if (textMatches) {
                                    const slideContent = textMatches.map(match => 
                                        match.replace(/<a:t[^>]*>([^<]*)<\/a:t>/, '$1')
                                    ).join(' ');
                                    slides.push(slideContent);
                                    slideText += `\n\nSlide ${slides.length}:\n${slideContent}`;
                                }
                                zipfile.readEntry();
                            });
                        });
                    } else {
                        zipfile.readEntry();
                    }
                });

                zipfile.on("end", () => {
                    resolve({
                        text: slideText.trim() || "[PowerPoint content - no extractable text found]",
                        metadata: {
                            slides: slides.length,
                            note: "Extracted text content from PowerPoint slides"
                        }
                    });
                });

                zipfile.readEntry();
            });
        });
    }

    /**
     * Parse plain text files
     */
    static parseText(buffer) {
        return {
            text: buffer.toString('utf8'),
            metadata: {
                encoding: 'utf8',
                size: buffer.length
            }
        };
    }

    /**
     * Parse HTML files
     */
    static parseHTML(buffer) {
        const html = buffer.toString('utf8');
        const root = parseHtml(html);
        
        // Extract text content, removing scripts and styles
        root.querySelectorAll('script, style').forEach(el => el.remove());
        const text = root.text;

        return {
            text: text,
            metadata: {
                title: root.querySelector('title')?.text || '',
                hasImages: root.querySelectorAll('img').length > 0,
                hasLinks: root.querySelectorAll('a').length > 0
            }
        };
    }

    /**
     * Parse RTF files (basic support)
     */
    static parseRTF(buffer) {
        const rtfContent = buffer.toString('utf8');
        // Basic RTF text extraction (removes most RTF control codes)
        const text = rtfContent
            .replace(/\\[a-z]+\d*\s*/g, '') // Remove RTF control words
            .replace(/[{}]/g, '') // Remove braces
            .replace(/\\\'/g, "'") // Handle escaped quotes
            .replace(/\\[\\{}]/g, '') // Handle escaped characters
            .trim();

        return {
            text: text,
            metadata: {
                note: "Basic RTF parsing - formatting information removed"
            }
        };
    }
}

/**
 * SharePoint MCP Server implementation
 * Provides tools and resources for interacting with Microsoft SharePoint via Microsoft Graph API
 */
class SharePointServer {
    server;
    accessToken = null;
    tokenExpiry = 0;
    constructor() {
        this.server = new Server({
            name: "sharepoint-mcp-server",
            version: "0.1.0",
        }, {
            capabilities: {
                tools: {},
                resources: {},
            },
        });
        this.setupHandlers();
        this.setupErrorHandling();
    }
    /**
     * Get access token for Microsoft Graph API
     */
    async getAccessToken() {
        if (this.accessToken && Date.now() < this.tokenExpiry) {
            return this.accessToken;
        }
        const tenantId = TENANT_ID;
        const clientId = CLIENT_ID;
        const clientSecret = CLIENT_SECRET;
        const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
        const params = new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            scope: "https://graph.microsoft.com/.default",
            grant_type: "client_credentials",
        });
        try {
            const response = await fetch(tokenUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded",
                },
                body: params,
            });
            if (!response.ok) {
                throw new Error(`Token request failed: ${response.status} ${response.statusText}`);
            }
            const data = await response.json();
            this.accessToken = data.access_token;
            this.tokenExpiry = Date.now() + (data.expires_in * 1000) - 60000; // Refresh 1 minute early
            return this.accessToken;
        }
        catch (error) {
            throw new Error(`Failed to get access token: ${error}`);
        }
    }
    /**
     * Make authenticated request to Microsoft Graph API
     */
    async graphRequest(endpoint, method = "GET", body) {
        const token = await this.getAccessToken();
        const url = `https://graph.microsoft.com/v1.0${endpoint}`;
        const headers = {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json",
        };
        const options = {
            method,
            headers,
        };
        if (body && method !== "GET") {
            options.body = JSON.stringify(body);
        }
        try {
            const response = await fetch(url, options);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Graph API request failed: ${response.status} ${response.statusText} - ${errorText}`);
            }
            return await response.json();
        }
        catch (error) {
            throw new Error(`Graph API request error: ${error}`);
        }
    }
    /**
     * Setup error handling for the server
     */
    setupErrorHandling() {
        this.server.onerror = (error) => console.error("[MCP Error]", error);
        process.on("SIGINT", async () => {
            await this.server.close();
            process.exit(0);
        });
    }
    /**
     * Setup all request handlers for tools and resources
     */
    setupHandlers() {
        this.setupToolHandlers();
        this.setupResourceHandlers();
    }
    /**
     * Setup tool handlers for SharePoint operations
     */
    setupToolHandlers() {
        this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
            tools: [
                {
                    name: "search_files",
                    description: "Search for files and documents in SharePoint using Microsoft Graph Search API. Can be scoped to specific site and folder to reduce noise.",
                    inputSchema: {
                        type: "object",
                        properties: {
                            query: {
                                type: "string",
                                description: "The search query string",
                            },
                            siteUrl: {
                                type: "string",
                                description: "Optional SharePoint site URL to scope search to (uses DEFAULT_SITE_URL if not provided)",
                            },
                            folderPath: {
                                type: "string",
                                description: "Optional folder path to scope search to (uses DEFAULT_FOLDER_PATH if not provided)",
                            },
                            limit: {
                                type: "number",
                                description: "Maximum number of results to return (default: 10)",
                                default: 10,
                            },
                        },
                        required: ["query"],
                    },
                },
                {
                    name: "list_sites",
                    description: "List SharePoint sites accessible to the application",
                    inputSchema: {
                        type: "object",
                        properties: {
                            search: {
                                type: "string",
                                description: "Optional search term to filter sites",
                            },
                        },
                    },
                },
                {
                    name: "get_site_info",
                    description: "Get detailed information about a specific SharePoint site",
                    inputSchema: {
                        type: "object",
                        properties: {
                            siteUrl: {
                                type: "string",
                                description: "The SharePoint site URL (e.g., https://tenant.sharepoint.com/sites/sitename)",
                            },
                        },
                        required: ["siteUrl"],
                    },
                },
                {
                    name: "list_site_drives",
                    description: "List document libraries (drives) in a SharePoint site",
                    inputSchema: {
                        type: "object",
                        properties: {
                            siteUrl: {
                                type: "string",
                                description: "The SharePoint site URL",
                            },
                        },
                        required: ["siteUrl"],
                    },
                },
                {
                    name: "list_drive_items",
                    description: "List files and folders in a SharePoint document library. Uses DEFAULT_SITE_URL and DEFAULT_FOLDER_PATH if configured to reduce noise.",
                    inputSchema: {
                        type: "object",
                        properties: {
                            siteUrl: {
                                type: "string",
                                description: "The SharePoint site URL (uses DEFAULT_SITE_URL if not provided)",
                            },
                            driveId: {
                                type: "string",
                                description: "The drive ID (optional, uses default drive if not specified)",
                            },
                            folderPath: {
                                type: "string",
                                description: "Optional folder path to list items from (uses DEFAULT_FOLDER_PATH if available, otherwise root)",
                            },
                        },
                        required: [],
                    },
                },
                {
                    name: "get_file_content",
                    description: "Get the content of a specific file from SharePoint. Supports multiple document formats including PDF, Word (doc/docx), Excel (xls/xlsx), PowerPoint (ppt/pptx), text files, HTML, RTF, and more. Uses DEFAULT_SITE_URL if configured.",
                    inputSchema: {
                        type: "object",
                        properties: {
                            siteUrl: {
                                type: "string",
                                description: "The SharePoint site URL (uses DEFAULT_SITE_URL if not provided)",
                            },
                            filePath: {
                                type: "string",
                                description: "The path to the file",
                            },
                            driveId: {
                                type: "string",
                                description: "The drive ID (optional, uses default drive if not specified)",
                            },
                            includeMetadata: {
                                type: "boolean",
                                description: "Whether to include document metadata in the response (default: true)",
                                default: true,
                            },
                        },
                        required: ["filePath"],
                    },
                },
            ],
        }));
        this.server.setRequestHandler(CallToolRequestSchema, async (request) => {
            try {
                switch (request.params.name) {
                    case "search_files":
                        return await this.handleSearchFiles(request.params.arguments);
                    case "list_sites":
                        return await this.handleListSites(request.params.arguments);
                    case "get_site_info":
                        return await this.handleGetSiteInfo(request.params.arguments);
                    case "list_site_drives":
                        return await this.handleListSiteDrives(request.params.arguments);
                    case "list_drive_items":
                        return await this.handleListDriveItems(request.params.arguments);
                    case "get_file_content":
                        return await this.handleGetFileContent(request.params.arguments);
                    default:
                        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
                }
            }
            catch (error) {
                const errorMessage = error instanceof Error ? error.message : String(error);
                throw new McpError(ErrorCode.InternalError, `SharePoint operation failed: ${errorMessage}`);
            }
        });
    }
    /**
     * Setup resource handlers for SharePoint resources
     */
    setupResourceHandlers() {
        this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
            try {
                const response = await this.graphRequest("/sites?$select=id,displayName,name,webUrl");
                const sites = response.value || [];
                return {
                    resources: sites.map((site) => ({
                        uri: `sharepoint://site/${site.id}`,
                        mimeType: "application/json",
                        name: site.displayName || site.name,
                        description: `SharePoint site: ${site.displayName || site.name} (${site.webUrl})`,
                    })),
                };
            }
            catch (error) {
                console.error("Error listing resources:", error);
                return { resources: [] };
            }
        });
        this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
            const url = new URL(request.params.uri);
            if (url.protocol === "sharepoint:" && url.pathname.startsWith("/site/")) {
                const siteId = url.pathname.replace("/site/", "");
                try {
                    const site = await this.graphRequest(`/sites/${siteId}`);
                    return {
                        contents: [{
                                uri: request.params.uri,
                                mimeType: "application/json",
                                text: JSON.stringify(site, null, 2),
                            }],
                    };
                }
                catch (error) {
                    throw new McpError(ErrorCode.InternalError, `Failed to read site resource: ${error}`);
                }
            }
            throw new McpError(ErrorCode.InvalidParams, `Unsupported resource URI: ${request.params.uri}`);
        });
    }
    /**
     * Extract site ID from SharePoint URL
     */
    async getSiteIdFromUrl(siteUrl) {
        try {
            const url = new URL(siteUrl);
            const hostname = url.hostname;
            const pathname = url.pathname;
            const response = await this.graphRequest(`/sites/${hostname}:${pathname}`);
            return response.id;
        }
        catch (error) {
            throw new Error(`Failed to get site ID from URL ${siteUrl}: ${error}`);
        }
    }
    /**
     * Handle search files tool request
     */
    async handleSearchFiles(args) {
        const query = args?.query;
        const siteUrl = args?.siteUrl || DEFAULT_SITE_URL;
        const folderPath = args?.folderPath || DEFAULT_FOLDER_PATH;
        const limit = args?.limit || 10;
        
        if (typeof query !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "Query parameter must be a string");
        }
        
        try {
            let searchQuery = query;
            
            // If we have site and/or folder constraints, enhance the search query
            if (siteUrl) {
                // Extract site name from URL for path constraint
                const url = new URL(siteUrl);
                const siteName = url.pathname.split('/').pop();
                searchQuery += ` path:"${siteName}"`;
                
                if (folderPath) {
                    searchQuery += ` AND path:"${folderPath}"`;
                }
            }
            
            const searchRequest = {
                requests: [{
                        entityTypes: ["driveItem"],
                        query: {
                            queryString: searchQuery,
                        },
                        size: limit,
                    }],
            };
            
            const searchResults = await this.graphRequest("/search/query", "POST", searchRequest);
            
            // If we have site filtering, further filter results to ensure they're from the correct site
            if (siteUrl && searchResults.value && searchResults.value[0]?.hitsContainers) {
                const siteId = await this.getSiteIdFromUrl(siteUrl);
                searchResults.value[0].hitsContainers = searchResults.value[0].hitsContainers.map(container => ({
                    ...container,
                    hits: container.hits?.filter(hit => 
                        hit.resource?.parentReference?.siteId === siteId ||
                        hit.resource?.webUrl?.includes(siteUrl)
                    ) || []
                }));
            }
            
            return {
                content: [{
                        type: "text",
                        text: JSON.stringify(searchResults, null, 2),
                    }],
            };
        }
        catch (error) {
            throw new Error(`Search failed: ${error}`);
        }
    }
    /**
     * Handle list sites tool request
     */
    async handleListSites(args) {
        const searchTerm = args?.search;
        try {
            let endpoint = "/sites?$select=id,displayName,name,webUrl,description";
            if (searchTerm) {
                endpoint += `&$filter=contains(displayName,'${searchTerm}')`;
            }
            const response = await this.graphRequest(endpoint);
            const sites = response.value || [];
            return {
                content: [{
                        type: "text",
                        text: JSON.stringify(sites, null, 2),
                    }],
            };
        }
        catch (error) {
            // Enhanced error handling for permission issues
            const errorMessage = error.toString();
            if (errorMessage.includes('Cannot enumerate sites') || errorMessage.includes('invalidRequest')) {
                const helpfulError = {
                    error: "Permission Error: Cannot enumerate sites",
                    message: "The Azure app registration lacks sufficient permissions to list all SharePoint sites.",
                    solution: "Add these Microsoft Graph Application permissions and grant admin consent:",
                    requiredPermissions: [
                        "Sites.Read.All - Read items in all site collections",
                        "Sites.ReadWrite.All - Read and write items in all site collections", 
                        "Directory.Read.All - Read directory data (may be required for site enumeration)"
                    ],
                    steps: [
                        "1. Go to Azure Portal → Azure Active Directory → App registrations",
                        "2. Find your app registration → API permissions",
                        "3. Add Microsoft Graph Application permissions: Sites.Read.All, Sites.ReadWrite.All, Directory.Read.All",
                        "4. Click 'Grant admin consent for your organization'",
                        "5. Wait a few minutes for permissions to propagate"
                    ],
                    workaround: "Use get_site_info tool with specific SharePoint site URLs instead of listing all sites",
                    originalError: errorMessage
                };
                return {
                    content: [{
                        type: "text",
                        text: JSON.stringify(helpfulError, null, 2),
                    }],
                };
            }
            throw new Error(`Failed to list sites: ${error}`);
        }
    }
    /**
     * Handle get site info tool request
     */
    async handleGetSiteInfo(args) {
        const siteUrl = args?.siteUrl;
        if (typeof siteUrl !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
        }
        try {
            const siteId = await this.getSiteIdFromUrl(siteUrl);
            const site = await this.graphRequest(`/sites/${siteId}?$expand=drive`);
            return {
                content: [{
                        type: "text",
                        text: JSON.stringify(site, null, 2),
                    }],
            };
        }
        catch (error) {
            throw new Error(`Failed to get site info: ${error}`);
        }
    }
    /**
     * Handle list site drives tool request
     */
    async handleListSiteDrives(args) {
        const siteUrl = args?.siteUrl;
        if (typeof siteUrl !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
        }
        try {
            const siteId = await this.getSiteIdFromUrl(siteUrl);
            const response = await this.graphRequest(`/sites/${siteId}/drives`);
            const drives = response.value || [];
            return {
                content: [{
                        type: "text",
                        text: JSON.stringify(drives, null, 2),
                    }],
            };
        }
        catch (error) {
            throw new Error(`Failed to list site drives: ${error}`);
        }
    }
    /**
     * Handle list drive items tool request
     */
    async handleListDriveItems(args) {
        const siteUrl = args?.siteUrl || DEFAULT_SITE_URL;
        const driveId = args?.driveId;
        const folderPath = args?.folderPath || DEFAULT_FOLDER_PATH;
        
        if (!siteUrl) {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be provided or DEFAULT_SITE_URL must be set");
        }
        if (typeof siteUrl !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be a string");
        }
        
        try {
            const siteId = await this.getSiteIdFromUrl(siteUrl);
            let endpoint;
            if (driveId) {
                if (folderPath) {
                    endpoint = `/sites/${siteId}/drives/${driveId}/root:/${folderPath}:/children`;
                }
                else {
                    endpoint = `/sites/${siteId}/drives/${driveId}/root/children`;
                }
            }
            else {
                if (folderPath) {
                    endpoint = `/sites/${siteId}/drive/root:/${folderPath}:/children`;
                }
                else {
                    endpoint = `/sites/${siteId}/drive/root/children`;
                }
            }
            const response = await this.graphRequest(endpoint);
            const items = response.value || [];
            return {
                content: [{
                        type: "text",
                        text: JSON.stringify(items, null, 2),
                    }],
            };
        }
        catch (error) {
            throw new Error(`Failed to list drive items: ${error}`);
        }
    }
    /**
     * Handle get file content tool request
     */
    async handleGetFileContent(args) {
        const siteUrl = args?.siteUrl || DEFAULT_SITE_URL;
        const filePath = args?.filePath;
        const driveId = args?.driveId;
        const includeMetadata = args?.includeMetadata !== false; // Default to true

        if (!siteUrl) {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl parameter must be provided or DEFAULT_SITE_URL must be set");
        }
        if (typeof siteUrl !== "string" || typeof filePath !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "siteUrl and filePath parameters must be strings");
        }

        try {
            const siteId = await this.getSiteIdFromUrl(siteUrl);
            const filename = filePath.split('/').pop() || '';
            
            // First, get file metadata to determine MIME type
            let metadataEndpoint;
            if (driveId) {
                metadataEndpoint = `/sites/${siteId}/drives/${driveId}/root:/${filePath}`;
            } else {
                metadataEndpoint = `/sites/${siteId}/drive/root:/${filePath}`;
            }

            const fileMetadata = await this.graphRequest(metadataEndpoint);
            const mimeType = fileMetadata.file?.mimeType || '';

            // Get file content
            let contentEndpoint;
            if (driveId) {
                contentEndpoint = `/sites/${siteId}/drives/${driveId}/root:/${filePath}:/content`;
            } else {
                contentEndpoint = `/sites/${siteId}/drive/root:/${filePath}:/content`;
            }

            const token = await this.getAccessToken();
            const response = await fetch(`https://graph.microsoft.com/v1.0${contentEndpoint}`, {
                headers: {
                    "Authorization": `Bearer ${token}`,
                },
            });

            if (!response.ok) {
                throw new Error(`Failed to get file content: ${response.status} ${response.statusText}`);
            }

            // Get file content as buffer for parsing
            const buffer = Buffer.from(await response.arrayBuffer());

            try {
                // Parse the document using our document parser
                const parseResult = await DocumentParser.parseDocument(buffer, filename, mimeType);
                
                let responseText = parseResult.text;
                
                if (includeMetadata && parseResult.metadata) {
                    responseText += `\n\n--- Document Metadata ---\n`;
                    responseText += `File: ${filename}\n`;
                    responseText += `MIME Type: ${mimeType}\n`;
                    responseText += `Size: ${fileMetadata.size} bytes\n`;
                    responseText += `Modified: ${fileMetadata.lastModifiedDateTime}\n`;
                    
                    if (parseResult.metadata) {
                        responseText += `Parser Metadata:\n${JSON.stringify(parseResult.metadata, null, 2)}`;
                    }
                }

                return {
                    content: [{
                        type: "text",
                        text: responseText,
                    }],
                };
            } catch (parseError) {
                // If parsing fails, fall back to treating as text if possible
                if (mimeType && mimeType.startsWith('text/')) {
                    const fallbackText = buffer.toString('utf8');
                    return {
                        content: [{
                            type: "text",
                            text: fallbackText + (includeMetadata ? `\n\n--- Document Metadata ---\nFile: ${filename}\nMIME Type: ${mimeType}\nSize: ${fileMetadata.size} bytes\nNote: Parsed as plain text due to parsing error: ${parseError.message}` : ''),
                        }],
                    };
                } else {
                    throw new Error(`Unable to parse ${filename}: ${parseError.message}. File format may not be supported or file may be corrupted.`);
                }
            }
        }
        catch (error) {
            throw new Error(`Failed to get file content: ${error}`);
        }
    }
    /**
     * Start the MCP server
     */
    async run() {
        const transport = new StdioServerTransport();
        await this.server.connect(transport);
        console.error("SharePoint MCP server running on stdio");
    }
}
/**
 * Main entry point
 */
const server = new SharePointServer();
server.run().catch((error) => {
    console.error("Failed to start SharePoint MCP server:", error);
    process.exit(1);
});
