#!/usr/bin/env node
/**
 * SharePoint MCP Server - SECURITY ENHANCED VERSION
 * 
 * Enforces strict access control to only DEFAULT_SITE_URL and DEFAULT_FOLDER_PATH
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
const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, DEFAULT_SITE_URL, DEFAULT_FOLDER_PATH, SEARCH_REGION } = process.env;

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error("Required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET");
}

// SECURITY: Enforce that DEFAULT_SITE_URL and DEFAULT_FOLDER_PATH are set
if (!DEFAULT_SITE_URL) {
    throw new Error("SECURITY: DEFAULT_SITE_URL must be set to enforce access restrictions");
}

if (!DEFAULT_FOLDER_PATH) {
    throw new Error("SECURITY: DEFAULT_FOLDER_PATH must be set to enforce access restrictions");
}

console.error(`[SECURITY] Access restricted to site: ${DEFAULT_SITE_URL}`);
console.error(`[SECURITY] Access restricted to folder: ${DEFAULT_FOLDER_PATH}`);

/**
 * Security validator to ensure all operations are within allowed boundaries
 */
class SecurityValidator {
    static validateSiteAccess(requestedSiteUrl) {
        if (!requestedSiteUrl) {
            return; // Will use DEFAULT_SITE_URL
        }
        
        // Normalize URLs for comparison
        const normalizeUrl = (url) => url.toLowerCase().replace(/\/$/, '');
        const allowedSite = normalizeUrl(DEFAULT_SITE_URL);
        const requestedSite = normalizeUrl(requestedSiteUrl);
        
        if (requestedSite !== allowedSite) {
            throw new Error(`SECURITY VIOLATION: Access denied to site '${requestedSiteUrl}'. Only '${DEFAULT_SITE_URL}' is allowed.`);
        }
    }
    
    static validateFolderAccess(requestedFolderPath) {
        if (!requestedFolderPath) {
            return; // Will use DEFAULT_FOLDER_PATH
        }
        
        // Normalize paths for comparison
        const normalizePath = (path) => path.replace(/^\/+|\/+$/g, '').toLowerCase();
        const allowedFolder = normalizePath(DEFAULT_FOLDER_PATH);
        const requestedFolder = normalizePath(requestedFolderPath);
        
        // Check if requested folder is the allowed folder or a subfolder of it
        if (requestedFolder !== allowedFolder && !requestedFolder.startsWith(allowedFolder + '/')) {
            throw new Error(`SECURITY VIOLATION: Access denied to folder '${requestedFolderPath}'. Only '${DEFAULT_FOLDER_PATH}' and its subfolders are allowed.`);
        }
    }
    
    static validateFileAccess(filePath) {
        if (!filePath) {
            throw new Error("File path is required");
        }
        
        // Normalize the file path
        const normalizedPath = filePath.replace(/^\/+/, '');
        const allowedFolderNormalized = DEFAULT_FOLDER_PATH.replace(/^\/+|\/+$/g, '');
        
        // Check if file is within the allowed folder
        if (!normalizedPath.toLowerCase().startsWith(allowedFolderNormalized.toLowerCase() + '/')) {
            throw new Error(`SECURITY VIOLATION: Access denied to file '${filePath}'. Only files within '${DEFAULT_FOLDER_PATH}' are allowed.`);
        }
    }
}

/**
 * Extract SharePoint tenant URL from a site URL
 */
function getSharePointTenantUrl(siteUrl) {
    if (!siteUrl) {
        throw new Error("Site URL is required to determine SharePoint tenant");
    }
    try {
        const url = new URL(siteUrl);
        return `${url.protocol}//${url.hostname}`;
    } catch (error) {
        throw new Error(`Invalid site URL format: ${siteUrl}`);
    }
}

// [DocumentParser class remains the same - including all the parsing methods]
class DocumentParser {
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
                    if (mimeType && mimeType.startsWith('text/')) {
                        return this.parseText(buffer);
                    }
                    throw new Error(`Unsupported file format: ${extension}`);
            }
        } catch (error) {
            throw new Error(`Failed to parse ${extension} file: ${error.message}`);
        }
    }

    static async parsePDF(buffer) {
        try {
            if (!Buffer.isBuffer(buffer)) {
                throw new Error('PDF data must be a Buffer');
            }
            
            if (buffer.length === 0) {
                throw new Error('PDF buffer is empty');
            }
            
            const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
            
            const loadingTask = pdfjs.getDocument({
                data: new Uint8Array(buffer),
                verbosity: 0
            });
            
            const pdf = await loadingTask.promise;
            let fullText = '';
            
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                try {
                    const page = await pdf.getPage(pageNum);
                    const textContent = await page.getTextContent();
                    
                    const pageText = textContent.items
                        .map(item => item.str)
                        .join(' ')
                        .trim();
                    
                    if (pageText) {
                        fullText += `\n\n--- Page ${pageNum} ---\n${pageText}`;
                    }
                } catch (pageError) {
                    console.warn(`Error extracting text from page ${pageNum}:`, pageError.message);
                    fullText += `\n\n--- Page ${pageNum} ---\n[Error extracting text from this page: ${pageError.message}]`;
                }
            }
            
            return {
                text: fullText.trim() || '[No text content found in PDF]',
                metadata: {
                    pages: pdf.numPages,
                    extractedBy: 'pdfjs-dist',
                    pdfVersion: pdf.pdfInfo?.PDFFormatVersion || 'unknown',
                    size: buffer.length
                }
            };
            
        } catch (error) {
            console.error('PDF parsing error:', error.message);
            throw new Error(`PDF parsing failed: ${error.message}. This may be due to an encrypted, corrupted, or unsupported PDF format.`);
        }
    }

    static async parseWord(buffer) {
        const result = await mammoth.extractRawText({ buffer });
        return {
            text: result.value,
            metadata: {
                messages: result.messages
            }
        };
    }

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

    static async parsePowerPoint(buffer) {
        return new Promise((resolve, reject) => {
            yauzl.fromBuffer(buffer, { lazyEntries: true }, (err, zipfile) => {
                if (err) {
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

    static parseText(buffer) {
        return {
            text: buffer.toString('utf8'),
            metadata: {
                encoding: 'utf8',
                size: buffer.length
            }
        };
    }

    static parseHTML(buffer) {
        const html = buffer.toString('utf8');
        const root = parseHtml(html);
        
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

    static parseRTF(buffer) {
        const rtfContent = buffer.toString('utf8');
        const text = rtfContent
            .replace(/\\[a-z]+\d*\s*/g, '')
            .replace(/[{}]/g, '')
            .replace(/\\\'/g, "'")
            .replace(/\\[\\{}]/g, '')
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
 * SharePoint MCP Server implementation with enforced security restrictions
 */
class SharePointServer {
    server;
    accessToken = null;
    tokenExpiry = 0;
    allowedSiteId = null; // Cache the allowed site ID

    constructor() {
        this.server = new Server({
            name: "sharepoint-mcp-server-secured",
            version: "0.1.0",
        }, {
            capabilities: {
                tools: {},
                resources: {},
            },
        });
        this.setupHandlers();
        this.setupErrorHandling();
        
        // Pre-validate and cache the allowed site ID
        this.initializeSecurity();
    }

    async initializeSecurity() {
        try {
            this.allowedSiteId = await this.getSiteIdFromUrl(DEFAULT_SITE_URL);
            console.error(`[SECURITY] Allowed site ID cached: ${this.allowedSiteId}`);
        } catch (error) {
            console.error(`[SECURITY] Failed to initialize security - could not get site ID: ${error.message}`);
            throw new Error(`Security initialization failed: ${error.message}`);
        }
    }

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
            this.tokenExpiry = Date.now() + (data.expires_in * 1000) - 60000;
            return this.accessToken;
        } catch (error) {
            throw new Error(`Failed to get access token: ${error}`);
        }
    }

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
        } catch (error) {
            throw new Error(`Graph API request error: ${error}`);
        }
    }

    setupErrorHandling() {
        this.server.onerror = (error) => console.error("[MCP Error]", error);
        process.on("SIGINT", async () => {
            await this.server.close();
            process.exit(0);
        });
    }

    setupHandlers() {
        this.setupToolHandlers();
        this.setupResourceHandlers();
    }

    setupToolHandlers() {
        this.server.setRequestHandler(ListToolsRequestSchema, async () => ({
            tools: [
                {
                    name: "search_files",
                    description: `Search for files and documents within the restricted folder '${DEFAULT_FOLDER_PATH}' on site '${DEFAULT_SITE_URL}'. Security: Access is strictly limited to this location.`,
                    inputSchema: {
                        type: "object",
                        properties: {
                            query: {
                                type: "string",
                                description: "The search query string",
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
                    name: "get_site_info",
                    description: `Get information about the restricted SharePoint site: ${DEFAULT_SITE_URL}`,
                    inputSchema: {
                        type: "object",
                        properties: {},
                    },
                },
                {
                    name: "list_drive_items",
                    description: `List files and folders within the restricted path '${DEFAULT_FOLDER_PATH}' and its subfolders. Security: Cannot access other locations.`,
                    inputSchema: {
                        type: "object",
                        properties: {
                            folderPath: {
                                type: "string",
                                description: `Optional subfolder path within '${DEFAULT_FOLDER_PATH}' (default: root of allowed folder)`,
                            },
                        },
                    },
                },
                {
                    name: "get_file_content",
                    description: `Get content of files within the restricted folder '${DEFAULT_FOLDER_PATH}'. Supports PDF, Word, Excel, PowerPoint, text files, and more. Security: Only files within allowed folder can be accessed.`,
                    inputSchema: {
                        type: "object",
                        properties: {
                            filePath: {
                                type: "string",
                                description: `Path to file within '${DEFAULT_FOLDER_PATH}'`,
                            },
                            includeMetadata: {
                                type: "boolean",
                                description: "Whether to include document metadata (default: true)",
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
                    case "get_site_info":
                        return await this.handleGetSiteInfo(request.params.arguments);
                    case "list_drive_items":
                        return await this.handleListDriveItems(request.params.arguments);
                    case "get_file_content":
                        return await this.handleGetFileContent(request.params.arguments);
                    default:
                        throw new McpError(ErrorCode.MethodNotFound, `Unknown tool: ${request.params.name}`);
                }
            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : String(error);
                throw new McpError(ErrorCode.InternalError, `SharePoint operation failed: ${errorMessage}`);
            }
        });
    }

    setupResourceHandlers() {
        // SECURITY: Only expose the allowed site as a resource
        this.server.setRequestHandler(ListResourcesRequestSchema, async () => {
            try {
                const site = await this.graphRequest(`/sites/${this.allowedSiteId}`);
                return {
                    resources: [{
                        uri: `sharepoint://site/${site.id}`,
                        mimeType: "application/json",
                        name: site.displayName || site.name,
                        description: `SharePoint site: ${site.displayName || site.name} (RESTRICTED ACCESS)`,
                    }],
                };
            } catch (error) {
                console.error("Error listing resources:", error);
                return { resources: [] };
            }
        });

        this.server.setRequestHandler(ReadResourceRequestSchema, async (request) => {
            const url = new URL(request.params.uri);
            if (url.protocol === "sharepoint:" && url.pathname.startsWith("/site/")) {
                const siteId = url.pathname.replace("/site/", "");
                
                // SECURITY: Only allow access to the allowed site
                if (siteId !== this.allowedSiteId) {
                    throw new McpError(ErrorCode.InvalidParams, `SECURITY VIOLATION: Access denied to site resource`);
                }

                try {
                    const site = await this.graphRequest(`/sites/${siteId}`);
                    return {
                        contents: [{
                            uri: request.params.uri,
                            mimeType: "application/json",
                            text: JSON.stringify(site, null, 2),
                        }],
                    };
                } catch (error) {
                    throw new McpError(ErrorCode.InternalError, `Failed to read site resource: ${error}`);
                }
            }
            throw new McpError(ErrorCode.InvalidParams, `Unsupported resource URI: ${request.params.uri}`);
        });
    }

    async getSiteIdFromUrl(siteUrl) {
        try {
            const url = new URL(siteUrl);
            const hostname = url.hostname;
            const pathname = url.pathname;
            const response = await this.graphRequest(`/sites/${hostname}:${pathname}`);
            return response.id;
        } catch (error) {
            throw new Error(`Failed to get site ID from URL ${siteUrl}: ${error}`);
        }
    }

    async handleSearchFiles(args) {
        const query = args?.query;
        const limit = args?.limit || 10;

        if (typeof query !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "Query parameter must be a string");
        }

        console.error(`[SECURITY] Search restricted to: ${DEFAULT_SITE_URL}/${DEFAULT_FOLDER_PATH}`);

        try {
            // SECURITY: Search is automatically restricted to allowed site and folder
            const searchRequest = {
                requests: [{
                    entityTypes: ["driveItem"],
                    query: {
                        queryString: `${query} AND path:"${DEFAULT_FOLDER_PATH}"`,
                    },
                    size: limit,
                    region: SEARCH_REGION || "NAM",
                }],
            };

            const searchResults = await this.graphRequest("/search/query", "POST", searchRequest);

            // SECURITY: Additional filtering to ensure only allowed site results
            if (searchResults.value && searchResults.value[0]?.hitsContainers) {
                searchResults.value[0].hitsContainers = searchResults.value[0].hitsContainers.map(container => ({
                    ...container,
                    hits: container.hits?.filter(hit => {
                        const isAllowedSite = hit.resource?.parentReference?.siteId === this.allowedSiteId ||
                                            hit.resource?.webUrl?.includes(DEFAULT_SITE_URL);
                        const isWithinAllowedFolder = hit.resource?.parentReference?.path?.includes(DEFAULT_FOLDER_PATH) ||
                                                    hit.resource?.webUrl?.includes(DEFAULT_FOLDER_PATH);
                        return isAllowedSite && isWithinAllowedFolder;
                    }) || []
                }));
            }

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify(searchResults, null, 2),
                }],
            };
        } catch (error) {
            throw new Error(`Search failed: ${error}`);
        }
    }

    async handleGetSiteInfo(args) {
        // SECURITY: Only return info for the allowed site
        try {
            const site = await this.graphRequest(`/sites/${this.allowedSiteId}?$expand=drive`);
            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        ...site,
                        securityNote: `Access restricted to: ${DEFAULT_SITE_URL}/${DEFAULT_FOLDER_PATH}`
                    }, null, 2),
                }],
            };
        } catch (error) {
            throw new Error(`Failed to get site info: ${error}`);
        }
    }

    async handleListDriveItems(args) {
        const requestedFolderPath = args?.folderPath;

        // SECURITY: Validate folder access
        const effectiveFolderPath = requestedFolderPath ? 
            `${DEFAULT_FOLDER_PATH}/${requestedFolderPath}`.replace(/\/+/g, '/') : 
            DEFAULT_FOLDER_PATH;

        SecurityValidator.validateFolderAccess(effectiveFolderPath);

        console.error(`[SECURITY] Listing items in: ${effectiveFolderPath}`);

        try {
            const endpoint = `/sites/${this.allowedSiteId}/drive/root:/${effectiveFolderPath}:/children`;
            const response = await this.graphRequest(endpoint);
            const items = response.value || [];

            return {
                content: [{
                    type: "text",
                    text: JSON.stringify({
                        restrictedPath: effectiveFolderPath,
                        securityNote: `Access restricted to ${DEFAULT_SITE_URL}/${DEFAULT_FOLDER_PATH} and subfolders`,
                        items: items
                    }, null, 2),
                }],
            };
        } catch (error) {
            throw new Error(`Failed to list drive items: ${error}`);
        }
    }

    async handleGetFileContent(args) {
        const filePath = args?.filePath;
        const includeMetadata = args?.includeMetadata !== false;

        if (typeof filePath !== "string") {
            throw new McpError(ErrorCode.InvalidParams, "filePath parameter must be a string");
        }

        // SECURITY: Validate file access
        const fullFilePath = filePath.startsWith(DEFAULT_FOLDER_PATH) ? 
            filePath : 
            `${DEFAULT_FOLDER_PATH}/${filePath}`.replace(/\/+/g, '/');

        SecurityValidator.validateFileAccess(fullFilePath);

        console.error(`[SECURITY] Accessing file: ${fullFilePath}`);

        try {
            const filename = fullFilePath.split('/').pop() || '';
            
            // Get file metadata
            const metadataEndpoint = `/sites/${this.allowedSiteId}/drive/root:/${fullFilePath}`;
            const fileMetadata = await this.graphRequest(metadataEndpoint);
            const mimeType = fileMetadata.file?.mimeType || '';

            // Get file content
            const contentEndpoint = `/sites/${this.allowedSiteId}/drive/root:/${fullFilePath}:/content`;
            const token = await this.getAccessToken();
            const response = await fetch(`https://graph.microsoft.com/v1.0${contentEndpoint}`, {
                headers: {
                    "Authorization": `Bearer ${token}`,
                },
            });

            if (!response.ok) {
                throw new Error(`Failed to get file content: ${response.status} ${response.statusText}`);
            }

            const arrayBuffer = await response.arrayBuffer();
            if (!arrayBuffer || arrayBuffer.byteLength === 0) {
                throw new Error('File content is empty or could not be retrieved');
            }
            const buffer = Buffer.from(arrayBuffer);

            try {
                const parseResult = await DocumentParser.parseDocument(buffer, filename, mimeType);
                
                let responseText = parseResult.text;
                
                if (includeMetadata && parseResult.metadata) {
                    responseText += `\n\n--- Document Metadata ---\n`;
                    responseText += `File: ${filename}\n`;
                    responseText += `Full Path: ${fullFilePath}\n`;
                    responseText += `MIME Type: ${mimeType}\n`;
                    responseText += `Size: ${fileMetadata.size} bytes\n`;
                    responseText += `Modified: ${fileMetadata.lastModifiedDateTime}\n`;
                    responseText += `Security: Access restricted to ${DEFAULT_FOLDER_PATH}\n`;
                    
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
                if (mimeType && mimeType.startsWith('text/')) {
                    const fallbackText = buffer.toString('utf8');
                    return {
                        content: [{
                            type: "text",
                            text: fallbackText + (includeMetadata ? 
                                `\n\n--- Document Metadata ---\nFile: ${filename}\nFull Path: ${fullFilePath}\nMIME Type: ${mimeType}\nSize: ${fileMetadata.size} bytes\nSecurity: Access restricted to ${DEFAULT_FOLDER_PATH}\nNote: Parsed as plain text due to parsing error: ${parseError.message}` : ''),
                        }],
                    };
                } else {
                    throw new Error(`Unable to parse ${filename}: ${parseError.message}. File format may not be supported or file may be corrupted.`);
                }
            }
        } catch (error) {
            throw new Error(`Failed to get file content: ${error}`);
        }
    }

    async run() {
        const transport = new StdioServerTransport();
        await this.server.connect(transport);
        console.error("SharePoint MCP server running on stdio with STRICT SECURITY ENFORCEMENT");
        console.error(`[SECURITY] Restricted to site: ${DEFAULT_SITE_URL}`);
        console.error(`[SECURITY] Restricted to folder: ${DEFAULT_FOLDER_PATH}`);
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