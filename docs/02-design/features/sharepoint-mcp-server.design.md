# Design: SharePoint MCP Server

**Feature Name**: SharePoint MCP Server
**Level**: Dynamic
**Type**: Core Infrastructure
**Created**: 2026-02-14
**Status**: Design
**Plan Reference**: `docs/01-plan/features/sharepoint-mcp-server.plan.md`

---

## 1. Architecture Overview

### 1.1 System Layers

```
┌─────────────────────────────────────────────────────────────────────┐
│                         Request Flow                                │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. AI Client (Claude Desktop)                                     │
│     │                                                               │
│     ├─ MCP Protocol Request (JSON-RPC)                             │
│     └─ Header: Authorization: Bearer {MCP_API_KEY}                 │
│                                                                     │
│  2. Cloudflare Worker (index.ts)                                   │
│     │                                                               │
│     ├─ API Key Validation Middleware                               │
│     └─ Streamable HTTP Transport                                   │
│                                                                     │
│  3. MCP Server (server.ts)                                         │
│     │                                                               │
│     ├─ Tool Registration & Routing                                 │
│     └─ Input Validation                                            │
│                                                                     │
│  4. Graph Client (graph-client.ts)                                 │
│     │                                                               │
│     ├─ Token Management (auth.ts)                                  │
│     │  ├─ Check KV Cache                                           │
│     │  ├─ Client Credentials Grant (if expired)                    │
│     │  └─ Store in KV (1 hour TTL)                                 │
│     │                                                               │
│     └─ HTTP Client for Graph API                                   │
│                                                                     │
│  5. Microsoft Graph API                                            │
│     │                                                               │
│     └─ SharePoint Online Data                                      │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

### 1.2 Module Structure

```typescript
// src/index.ts - Cloudflare Worker Entry Point
export default {
  async fetch(request: Request, env: Env): Promise<Response>
}

// src/server.ts - MCP Server Instance
export function createMCPServer(env: Env): Server

// src/auth.ts - Authentication Management
export class AuthManager {
  async getAccessToken(): Promise<string>
  async refreshToken(): Promise<void>
}

// src/graph-client.ts - Graph API Client
export class GraphClient {
  async get<T>(endpoint: string): Promise<T>
  async post<T>(endpoint: string, body: any): Promise<T>
  async patch<T>(endpoint: string, body: any): Promise<T>
  async delete(endpoint: string): Promise<void>
}

// src/tools/documents.ts - Document Tools
export const searchDocumentsTool: Tool
export const getFileContentTool: Tool
export const listFilesTool: Tool

// src/tools/lists.ts - SharePoint List Tools
export const getListItemsTool: Tool
export const createListItemTool: Tool
export const updateListItemTool: Tool
export const deleteListItemTool: Tool
```

---

## 2. Data Models

### 2.1 Environment & Configuration

```typescript
// src/types/env.ts
export interface Env {
  // Secrets
  MCP_API_KEY: string;
  AZURE_CLIENT_ID: string;
  AZURE_CLIENT_SECRET: string;
  AZURE_TENANT_ID: string;
  SHAREPOINT_SITE_URL: string;

  // KV Namespace
  GRAPH_TOKEN_CACHE: KVNamespace;
}

export interface GraphTokenCache {
  access_token: string;
  expires_at: number; // Unix timestamp
  token_type: "Bearer";
}
```

### 2.2 MCP Tool Schemas

#### Tool: `search_documents`

```typescript
// Input Schema
export interface SearchDocumentsInput {
  query: string;              // Required: Search keyword
  site_id?: string;           // Optional: Specific site ID
  file_type?: string;         // Optional: Filter by extension (docx, xlsx, pdf)
  max_results?: number;       // Optional: Default 10, max 50
}

// Output Schema
export interface SearchDocumentsOutput {
  success: boolean;
  documents: DocumentMetadata[];
  total_count: number;
  query: string;
}

export interface DocumentMetadata {
  id: string;                 // driveItem ID
  name: string;               // File name
  web_url: string;            // SharePoint web URL
  created_datetime: string;   // ISO 8601
  last_modified_datetime: string;
  size: number;               // Bytes
  file_type: string;          // Extension
  created_by: string;         // User display name
  modified_by: string;
  parent_path: string;        // Folder path
}
```

#### Tool: `get_file_content`

```typescript
// Input Schema
export interface GetFileContentInput {
  file_id: string;            // Required: driveItem ID
  site_id?: string;           // Optional: Site ID
}

// Output Schema
export interface GetFileContentOutput {
  success: boolean;
  file_id: string;
  file_name: string;
  content_type: string;       // MIME type
  content: string;            // Text content (UTF-8)
  size: number;               // Bytes
  encoding: string;           // "utf-8"
}
```

#### Tool: `list_files`

```typescript
// Input Schema
export interface ListFilesInput {
  folder_path?: string;       // Optional: Default root "/"
  site_id?: string;           // Optional: Site ID
}

// Output Schema
export interface ListFilesOutput {
  success: boolean;
  folder_path: string;
  items: FileSystemItem[];
  count: number;
}

export interface FileSystemItem {
  id: string;
  name: string;
  type: "file" | "folder";
  size?: number;              // Only for files
  last_modified: string;      // ISO 8601
  web_url: string;
}
```

#### Tool: `get_list_items`

```typescript
// Input Schema
export interface GetListItemsInput {
  list_name: string;          // Required: List name or ID
  site_id?: string;           // Optional: Site ID
  filter?: string;            // Optional: OData filter query
  select?: string[];          // Optional: Fields to return
  top?: number;               // Optional: Default 50, max 100
}

// Output Schema
export interface GetListItemsOutput {
  success: boolean;
  list_name: string;
  items: ListItem[];
  count: number;
  has_more: boolean;
}

export interface ListItem {
  id: string;
  fields: Record<string, any>; // Dynamic field values
  created_datetime: string;
  last_modified_datetime: string;
  created_by: string;
  modified_by: string;
}
```

#### Tool: `create_list_item`

```typescript
// Input Schema
export interface CreateListItemInput {
  list_name: string;          // Required: List name or ID
  fields: Record<string, any>; // Required: Field values to set
  site_id?: string;           // Optional: Site ID
}

// Output Schema
export interface CreateListItemOutput {
  success: boolean;
  item_id: string;
  fields: Record<string, any>;
  created_datetime: string;
  web_url: string;
}
```

#### Tool: `update_list_item`

```typescript
// Input Schema
export interface UpdateListItemInput {
  list_name: string;          // Required: List name or ID
  item_id: string;            // Required: Item ID to update
  fields: Record<string, any>; // Required: Fields to update
  site_id?: string;           // Optional: Site ID
}

// Output Schema
export interface UpdateListItemOutput {
  success: boolean;
  item_id: string;
  fields: Record<string, any>; // Updated values
  last_modified_datetime: string;
}
```

#### Tool: `delete_list_item`

```typescript
// Input Schema
export interface DeleteListItemInput {
  list_name: string;          // Required: List name or ID
  item_id: string;            // Required: Item ID to delete
  site_id?: string;           // Optional: Site ID
}

// Output Schema
export interface DeleteListItemOutput {
  success: boolean;
  item_id: string;
  deleted_at: string;         // ISO 8601 timestamp
}
```

### 2.3 Microsoft Graph API Models

```typescript
// Graph API Response Types
export interface GraphDriveItem {
  id: string;
  name: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  size: number;
  file?: { mimeType: string };
  folder?: { childCount: number };
  createdBy: { user: { displayName: string } };
  lastModifiedBy: { user: { displayName: string } };
  parentReference: { path: string };
}

export interface GraphListItem {
  id: string;
  fields: Record<string, any>;
  createdDateTime: string;
  lastModifiedDateTime: string;
  createdBy: { user: { displayName: string } };
  lastModifiedBy: { user: { displayName: string } };
}

export interface GraphSearchResponse {
  value: GraphDriveItem[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
}

export interface GraphListResponse {
  value: GraphListItem[];
  "@odata.count"?: number;
  "@odata.nextLink"?: string;
}
```

### 2.4 Error Models

```typescript
export interface MCPError {
  code: string;
  message: string;
  details?: any;
}

export enum ErrorCode {
  // Authentication Errors (1xxx)
  INVALID_API_KEY = "1001",
  GRAPH_TOKEN_FAILED = "1002",
  UNAUTHORIZED = "1003",

  // Validation Errors (2xxx)
  INVALID_INPUT = "2001",
  MISSING_REQUIRED_FIELD = "2002",
  INVALID_FIELD_TYPE = "2003",

  // Graph API Errors (3xxx)
  GRAPH_API_ERROR = "3000",
  RESOURCE_NOT_FOUND = "3001",
  RATE_LIMIT_EXCEEDED = "3002",
  PERMISSION_DENIED = "3003",

  // Server Errors (5xxx)
  INTERNAL_ERROR = "5000",
  KV_STORAGE_ERROR = "5001",
}
```

---

## 3. Authentication Design

### 3.1 API Key Validation (Layer 1)

```typescript
// src/middleware/auth.ts
export async function validateAPIKey(
  request: Request,
  env: Env
): Promise<boolean> {
  const authHeader = request.headers.get("Authorization");

  if (!authHeader) {
    return false;
  }

  // Expected format: "Bearer {API_KEY}"
  const [scheme, token] = authHeader.split(" ");

  if (scheme !== "Bearer" || !token) {
    return false;
  }

  // Constant-time comparison to prevent timing attacks
  return safeCompare(token, env.MCP_API_KEY);
}

function safeCompare(a: string, b: string): boolean {
  if (a.length !== b.length) {
    return false;
  }

  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }

  return result === 0;
}
```

### 3.2 Client Credentials Flow (Layer 2)

```typescript
// src/auth.ts
export class AuthManager {
  private env: Env;
  private tokenCacheKey = "graph_access_token";

  constructor(env: Env) {
    this.env = env;
  }

  /**
   * Get valid access token (from cache or new)
   */
  async getAccessToken(): Promise<string> {
    // Check KV cache first
    const cached = await this.getCachedToken();

    if (cached && !this.isTokenExpired(cached)) {
      return cached.access_token;
    }

    // Request new token
    const token = await this.requestNewToken();

    // Cache in KV
    await this.cacheToken(token);

    return token.access_token;
  }

  /**
   * Request new token from Azure AD
   */
  private async requestNewToken(): Promise<GraphTokenCache> {
    const tokenEndpoint = `https://login.microsoftonline.com/${this.env.AZURE_TENANT_ID}/oauth2/v2.0/token`;

    const params = new URLSearchParams({
      client_id: this.env.AZURE_CLIENT_ID,
      client_secret: this.env.AZURE_CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials",
    });

    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params,
    });

    if (!response.ok) {
      throw new Error(`Token request failed: ${response.status}`);
    }

    const data = await response.json();

    return {
      access_token: data.access_token,
      expires_at: Date.now() + (data.expires_in * 1000) - 300000, // 5 min buffer
      token_type: "Bearer",
    };
  }

  /**
   * Get token from KV cache
   */
  private async getCachedToken(): Promise<GraphTokenCache | null> {
    const cached = await this.env.GRAPH_TOKEN_CACHE.get(
      this.tokenCacheKey,
      "json"
    );

    return cached as GraphTokenCache | null;
  }

  /**
   * Store token in KV cache
   */
  private async cacheToken(token: GraphTokenCache): Promise<void> {
    const ttl = Math.floor((token.expires_at - Date.now()) / 1000);

    await this.env.GRAPH_TOKEN_CACHE.put(
      this.tokenCacheKey,
      JSON.stringify(token),
      { expirationTtl: ttl }
    );
  }

  /**
   * Check if token is expired
   */
  private isTokenExpired(token: GraphTokenCache): boolean {
    return Date.now() >= token.expires_at;
  }
}
```

---

## 4. Microsoft Graph API Integration

### 4.1 Graph Client Implementation

```typescript
// src/graph-client.ts
export class GraphClient {
  private baseUrl = "https://graph.microsoft.com/v1.0";
  private authManager: AuthManager;

  constructor(env: Env) {
    this.authManager = new AuthManager(env);
  }

  /**
   * Generic GET request
   */
  async get<T>(endpoint: string): Promise<T> {
    const token = await this.authManager.getAccessToken();

    const response = await fetch(`${this.baseUrl}${endpoint}`, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    });

    return this.handleResponse<T>(response);
  }

  /**
   * Generic POST request
   */
  async post<T>(endpoint: string, body: any): Promise<T> {
    const token = await this.authManager.getAccessToken();

    const response = await fetch(`${this.baseUrl}${endpoint}`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    return this.handleResponse<T>(response);
  }

  /**
   * Generic PATCH request
   */
  async patch<T>(endpoint: string, body: any): Promise<T> {
    const token = await this.authManager.getAccessToken();

    const response = await fetch(`${this.baseUrl}${endpoint}`, {
      method: "PATCH",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    return this.handleResponse<T>(response);
  }

  /**
   * Generic DELETE request
   */
  async delete(endpoint: string): Promise<void> {
    const token = await this.authManager.getAccessToken();

    const response = await fetch(`${this.baseUrl}${endpoint}`, {
      method: "DELETE",
      headers: {
        "Authorization": `Bearer ${token}`,
      },
    });

    if (!response.ok && response.status !== 204) {
      throw await this.handleError(response);
    }
  }

  /**
   * Handle successful response
   */
  private async handleResponse<T>(response: Response): Promise<T> {
    if (!response.ok) {
      throw await this.handleError(response);
    }

    // Some Graph API responses return 204 No Content
    if (response.status === 204) {
      return {} as T;
    }

    return response.json();
  }

  /**
   * Handle error response
   */
  private async handleError(response: Response): Promise<Error> {
    let errorData: any;

    try {
      errorData = await response.json();
    } catch {
      errorData = { message: response.statusText };
    }

    const error: MCPError = {
      code: this.mapStatusToErrorCode(response.status),
      message: errorData.error?.message || errorData.message || "Graph API error",
      details: errorData,
    };

    return new Error(JSON.stringify(error));
  }

  /**
   * Map HTTP status to error code
   */
  private mapStatusToErrorCode(status: number): string {
    switch (status) {
      case 401:
      case 403:
        return ErrorCode.PERMISSION_DENIED;
      case 404:
        return ErrorCode.RESOURCE_NOT_FOUND;
      case 429:
        return ErrorCode.RATE_LIMIT_EXCEEDED;
      default:
        return ErrorCode.GRAPH_API_ERROR;
    }
  }
}
```

### 4.2 API Endpoint Mapping

```typescript
// Graph API Endpoints Used
export const GRAPH_ENDPOINTS = {
  // Documents
  SEARCH_DOCUMENTS: (siteId: string, query: string) =>
    `/sites/${siteId}/drive/root/search(q='${encodeURIComponent(query)}')`,

  GET_FILE_CONTENT: (siteId: string, itemId: string) =>
    `/sites/${siteId}/drive/items/${itemId}/content`,

  LIST_FILES: (siteId: string, path: string) =>
    `/sites/${siteId}/drive/root:${path}:/children`,

  // Lists
  GET_LIST_ITEMS: (siteId: string, listId: string) =>
    `/sites/${siteId}/lists/${listId}/items`,

  CREATE_LIST_ITEM: (siteId: string, listId: string) =>
    `/sites/${siteId}/lists/${listId}/items`,

  UPDATE_LIST_ITEM: (siteId: string, listId: string, itemId: string) =>
    `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,

  DELETE_LIST_ITEM: (siteId: string, listId: string, itemId: string) =>
    `/sites/${siteId}/lists/${listId}/items/${itemId}`,

  // Utility
  GET_SITE_BY_URL: (siteUrl: string) =>
    `/sites/${new URL(siteUrl).hostname}:${new URL(siteUrl).pathname}`,
};
```

---

## 5. MCP Server Implementation

### 5.1 Server Setup

```typescript
// src/server.ts
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamable.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema
} from "@modelcontextprotocol/sdk/types.js";

export function createMCPServer(env: Env): Server {
  const server = new Server(
    {
      name: "sharepoint-mcp-server",
      version: "1.0.0",
    },
    {
      capabilities: {
        tools: {},
      },
    }
  );

  // Register tool list handler
  server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
      tools: [
        // Document Tools
        {
          name: "search_documents",
          description: "Search SharePoint documents by keyword with optional filters",
          inputSchema: {
            type: "object",
            properties: {
              query: { type: "string", description: "Search keyword" },
              site_id: { type: "string", description: "Optional site ID" },
              file_type: { type: "string", description: "File extension filter" },
              max_results: { type: "number", description: "Max results (default 10)" },
            },
            required: ["query"],
          },
        },
        {
          name: "get_file_content",
          description: "Download and extract text content from a SharePoint file",
          inputSchema: {
            type: "object",
            properties: {
              file_id: { type: "string", description: "DriveItem ID" },
              site_id: { type: "string", description: "Optional site ID" },
            },
            required: ["file_id"],
          },
        },
        {
          name: "list_files",
          description: "List files and folders in a SharePoint directory",
          inputSchema: {
            type: "object",
            properties: {
              folder_path: { type: "string", description: "Folder path (default root)" },
              site_id: { type: "string", description: "Optional site ID" },
            },
            required: [],
          },
        },

        // List Tools
        {
          name: "get_list_items",
          description: "Query SharePoint list items with OData filtering",
          inputSchema: {
            type: "object",
            properties: {
              list_name: { type: "string", description: "List name or ID" },
              site_id: { type: "string", description: "Optional site ID" },
              filter: { type: "string", description: "OData filter query" },
              select: { type: "array", items: { type: "string" }, description: "Fields to return" },
              top: { type: "number", description: "Max items (default 50)" },
            },
            required: ["list_name"],
          },
        },
        {
          name: "create_list_item",
          description: "Create a new item in a SharePoint list",
          inputSchema: {
            type: "object",
            properties: {
              list_name: { type: "string", description: "List name or ID" },
              fields: { type: "object", description: "Field values to set" },
              site_id: { type: "string", description: "Optional site ID" },
            },
            required: ["list_name", "fields"],
          },
        },
        {
          name: "update_list_item",
          description: "Update an existing SharePoint list item",
          inputSchema: {
            type: "object",
            properties: {
              list_name: { type: "string", description: "List name or ID" },
              item_id: { type: "string", description: "Item ID to update" },
              fields: { type: "object", description: "Fields to update" },
              site_id: { type: "string", description: "Optional site ID" },
            },
            required: ["list_name", "item_id", "fields"],
          },
        },
        {
          name: "delete_list_item",
          description: "Delete a SharePoint list item",
          inputSchema: {
            type: "object",
            properties: {
              list_name: { type: "string", description: "List name or ID" },
              item_id: { type: "string", description: "Item ID to delete" },
              site_id: { type: "string", description: "Optional site ID" },
            },
            required: ["list_name", "item_id"],
          },
        },
      ],
    };
  });

  // Register tool call handler
  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: args } = request.params;

    try {
      const graphClient = new GraphClient(env);

      switch (name) {
        case "search_documents":
          return await handleSearchDocuments(graphClient, env, args);

        case "get_file_content":
          return await handleGetFileContent(graphClient, env, args);

        case "list_files":
          return await handleListFiles(graphClient, env, args);

        case "get_list_items":
          return await handleGetListItems(graphClient, env, args);

        case "create_list_item":
          return await handleCreateListItem(graphClient, env, args);

        case "update_list_item":
          return await handleUpdateListItem(graphClient, env, args);

        case "delete_list_item":
          return await handleDeleteListItem(graphClient, env, args);

        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              success: false,
              error: error instanceof Error ? error.message : String(error),
            }),
          },
        ],
      };
    }
  });

  return server;
}
```

### 5.2 Entry Point (Cloudflare Worker)

```typescript
// src/index.ts
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamable.js";
import { createMCPServer } from "./server";
import { validateAPIKey } from "./middleware/auth";

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    // CORS headers for browser clients
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    };

    // Handle preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders });
    }

    // Validate API Key
    const isValid = await validateAPIKey(request, env);

    if (!isValid) {
      return new Response(
        JSON.stringify({ error: "Invalid or missing API key" }),
        {
          status: 401,
          headers: {
            "Content-Type": "application/json",
            ...corsHeaders
          }
        }
      );
    }

    // Create MCP server
    const server = createMCPServer(env);

    // Create Streamable HTTP transport
    const transport = new StreamableHTTPServerTransport({
      endpoint: "/mcp",
    });

    // Connect server to transport
    await server.connect(transport);

    // Handle request
    const response = await transport.handleRequest(request);

    // Add CORS headers to response
    const headers = new Headers(response.headers);
    Object.entries(corsHeaders).forEach(([key, value]) => {
      headers.set(key, value);
    });

    return new Response(response.body, {
      status: response.status,
      headers,
    });
  },
};
```

---

## 6. Tool Handlers Implementation

### 6.1 Document Tools

```typescript
// src/tools/documents.ts

/**
 * Search SharePoint documents
 */
export async function handleSearchDocuments(
  client: GraphClient,
  env: Env,
  args: SearchDocumentsInput
): Promise<ToolResponse> {
  // Validate input
  if (!args.query || args.query.trim() === "") {
    throw new Error("Query parameter is required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);
  const maxResults = Math.min(args.max_results || 10, 50);

  // Build endpoint
  let endpoint = GRAPH_ENDPOINTS.SEARCH_DOCUMENTS(siteId, args.query);

  // Add file type filter if specified
  if (args.file_type) {
    endpoint += `&$filter=endswith(name,'${args.file_type}')`;
  }

  endpoint += `&$top=${maxResults}`;

  // Execute search
  const response = await client.get<GraphSearchResponse>(endpoint);

  // Transform to output format
  const documents: DocumentMetadata[] = response.value.map(item => ({
    id: item.id,
    name: item.name,
    web_url: item.webUrl,
    created_datetime: item.createdDateTime,
    last_modified_datetime: item.lastModifiedDateTime,
    size: item.size,
    file_type: item.name.split('.').pop() || "",
    created_by: item.createdBy.user.displayName,
    modified_by: item.lastModifiedBy.user.displayName,
    parent_path: item.parentReference.path,
  }));

  const result: SearchDocumentsOutput = {
    success: true,
    documents,
    total_count: response["@odata.count"] || documents.length,
    query: args.query,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Get file content
 */
export async function handleGetFileContent(
  client: GraphClient,
  env: Env,
  args: GetFileContentInput
): Promise<ToolResponse> {
  if (!args.file_id) {
    throw new Error("file_id is required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);

  // Get file metadata first
  const metadata = await client.get<GraphDriveItem>(
    `/sites/${siteId}/drive/items/${args.file_id}`
  );

  // Download file content
  const contentEndpoint = GRAPH_ENDPOINTS.GET_FILE_CONTENT(siteId, args.file_id);
  const contentResponse = await fetch(
    `https://graph.microsoft.com/v1.0${contentEndpoint}`,
    {
      headers: {
        Authorization: `Bearer ${await client["authManager"].getAccessToken()}`,
      },
    }
  );

  if (!contentResponse.ok) {
    throw new Error(`Failed to download file: ${contentResponse.status}`);
  }

  const content = await contentResponse.text();

  const result: GetFileContentOutput = {
    success: true,
    file_id: args.file_id,
    file_name: metadata.name,
    content_type: metadata.file?.mimeType || "application/octet-stream",
    content,
    size: metadata.size,
    encoding: "utf-8",
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * List files in folder
 */
export async function handleListFiles(
  client: GraphClient,
  env: Env,
  args: ListFilesInput
): Promise<ToolResponse> {
  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);
  const folderPath = args.folder_path || "/";

  const endpoint = GRAPH_ENDPOINTS.LIST_FILES(siteId, folderPath);
  const response = await client.get<GraphSearchResponse>(endpoint);

  const items: FileSystemItem[] = response.value.map(item => ({
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : "file",
    size: item.size,
    last_modified: item.lastModifiedDateTime,
    web_url: item.webUrl,
  }));

  const result: ListFilesOutput = {
    success: true,
    folder_path: folderPath,
    items,
    count: items.length,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Utility: Get site ID from URL
 */
async function getSiteIdFromUrl(client: GraphClient, siteUrl: string): Promise<string> {
  const endpoint = GRAPH_ENDPOINTS.GET_SITE_BY_URL(siteUrl);
  const site = await client.get<{ id: string }>(endpoint);
  return site.id;
}
```

### 6.2 SharePoint List Tools

```typescript
// src/tools/lists.ts

/**
 * Get list items with filtering
 */
export async function handleGetListItems(
  client: GraphClient,
  env: Env,
  args: GetListItemsInput
): Promise<ToolResponse> {
  if (!args.list_name) {
    throw new Error("list_name is required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);
  const top = Math.min(args.top || 50, 100);

  // Build query parameters
  const params = new URLSearchParams();
  params.set("$expand", "fields");
  params.set("$top", String(top));

  if (args.filter) {
    params.set("$filter", args.filter);
  }

  if (args.select && args.select.length > 0) {
    params.set("$select", args.select.join(","));
  }

  const endpoint = `${GRAPH_ENDPOINTS.GET_LIST_ITEMS(siteId, args.list_name)}?${params}`;
  const response = await client.get<GraphListResponse>(endpoint);

  const items: ListItem[] = response.value.map(item => ({
    id: item.id,
    fields: item.fields,
    created_datetime: item.createdDateTime,
    last_modified_datetime: item.lastModifiedDateTime,
    created_by: item.createdBy.user.displayName,
    modified_by: item.lastModifiedBy.user.displayName,
  }));

  const result: GetListItemsOutput = {
    success: true,
    list_name: args.list_name,
    items,
    count: items.length,
    has_more: !!response["@odata.nextLink"],
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Create list item
 */
export async function handleCreateListItem(
  client: GraphClient,
  env: Env,
  args: CreateListItemInput
): Promise<ToolResponse> {
  if (!args.list_name || !args.fields) {
    throw new Error("list_name and fields are required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);

  const endpoint = GRAPH_ENDPOINTS.CREATE_LIST_ITEM(siteId, args.list_name);
  const body = { fields: args.fields };

  const response = await client.post<GraphListItem>(endpoint, body);

  const result: CreateListItemOutput = {
    success: true,
    item_id: response.id,
    fields: response.fields,
    created_datetime: response.createdDateTime,
    web_url: `${env.SHAREPOINT_SITE_URL}/Lists/${args.list_name}/DispForm.aspx?ID=${response.id}`,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Update list item
 */
export async function handleUpdateListItem(
  client: GraphClient,
  env: Env,
  args: UpdateListItemInput
): Promise<ToolResponse> {
  if (!args.list_name || !args.item_id || !args.fields) {
    throw new Error("list_name, item_id, and fields are required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);

  const endpoint = GRAPH_ENDPOINTS.UPDATE_LIST_ITEM(siteId, args.list_name, args.item_id);
  const response = await client.patch<Record<string, any>>(endpoint, args.fields);

  const result: UpdateListItemOutput = {
    success: true,
    item_id: args.item_id,
    fields: response,
    last_modified_datetime: new Date().toISOString(),
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Delete list item
 */
export async function handleDeleteListItem(
  client: GraphClient,
  env: Env,
  args: DeleteListItemInput
): Promise<ToolResponse> {
  if (!args.list_name || !args.item_id) {
    throw new Error("list_name and item_id are required");
  }

  const siteId = args.site_id || await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL);

  const endpoint = GRAPH_ENDPOINTS.DELETE_LIST_ITEM(siteId, args.list_name, args.item_id);
  await client.delete(endpoint);

  const result: DeleteListItemOutput = {
    success: true,
    item_id: args.item_id,
    deleted_at: new Date().toISOString(),
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}
```

---

## 7. Error Handling Strategy

### 7.1 Error Handling Flow

```typescript
// Error handling at each layer

// Layer 1: API Key Validation
if (!validateAPIKey(request, env)) {
  return new Response("Unauthorized", { status: 401 });
}

// Layer 2: Input Validation
try {
  validateToolInput(args, schema);
} catch (error) {
  throw new MCPError(ErrorCode.INVALID_INPUT, error.message);
}

// Layer 3: Graph API Errors
try {
  await graphClient.get(endpoint);
} catch (error) {
  if (error.code === "404") {
    throw new MCPError(ErrorCode.RESOURCE_NOT_FOUND, "Resource not found");
  }
  throw error;
}

// Layer 4: Response Formatting
return {
  content: [{
    type: "text",
    text: JSON.stringify({ success: false, error: error.message })
  }]
};
```

### 7.2 Retry Strategy

```typescript
// Retry with exponential backoff for rate limiting
export async function retryWithBackoff<T>(
  fn: () => Promise<T>,
  maxRetries: number = 3
): Promise<T> {
  let lastError: Error;

  for (let i = 0; i < maxRetries; i++) {
    try {
      return await fn();
    } catch (error) {
      lastError = error as Error;

      // Check if error is retryable (429 or 5xx)
      const errorData = JSON.parse(lastError.message);
      if (errorData.code === ErrorCode.RATE_LIMIT_EXCEEDED) {
        const delay = Math.pow(2, i) * 1000; // 1s, 2s, 4s
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }

      // Non-retryable error
      throw error;
    }
  }

  throw lastError!;
}
```

---

## 8. Configuration Files

### 8.1 wrangler.jsonc

```jsonc
{
  "name": "sharepoint-mcp-server",
  "main": "src/index.ts",
  "compatibility_date": "2024-01-01",

  // KV Namespace binding
  "kv_namespaces": [
    {
      "binding": "GRAPH_TOKEN_CACHE",
      "id": "your-kv-namespace-id",
      "preview_id": "your-preview-kv-namespace-id"
    }
  ],

  // Environment variables (non-secret)
  "vars": {
    "ENVIRONMENT": "production"
  },

  // Note: Secrets are set via wrangler secret put command
  // - MCP_API_KEY
  // - AZURE_CLIENT_ID
  // - AZURE_CLIENT_SECRET
  // - AZURE_TENANT_ID
  // - SHAREPOINT_SITE_URL
}
```

### 8.2 package.json

```json
{
  "name": "sharepoint-mcp-server",
  "version": "1.0.0",
  "type": "module",
  "scripts": {
    "dev": "wrangler dev",
    "deploy": "wrangler deploy",
    "tail": "wrangler tail",
    "cf-typegen": "wrangler types"
  },
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.0.0"
  },
  "devDependencies": {
    "@cloudflare/workers-types": "^4.20240117.0",
    "typescript": "^5.3.3",
    "wrangler": "^3.22.1"
  }
}
```

### 8.3 tsconfig.json

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ES2020",
    "lib": ["ES2020"],
    "moduleResolution": "node",
    "types": ["@cloudflare/workers-types"],
    "resolveJsonModule": true,
    "allowJs": true,
    "checkJs": false,
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "isolatedModules": true,
    "noEmit": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules"]
}
```

---

## 9. Testing Strategy

### 9.1 Manual Test Cases

**Authentication Tests**:
1. ✅ API Key validation with valid key
2. ❌ API Key validation with invalid key
3. ❌ API Key validation with missing header
4. ✅ Graph API token acquisition
5. ✅ Token caching and reuse

**Document Tools Tests**:
1. Search with keyword only
2. Search with file type filter
3. Search with max_results parameter
4. Get file content for text file
5. Get file content for non-existent file (error case)
6. List files in root directory
7. List files in specific folder

**List Tools Tests**:
1. Get all list items
2. Get list items with OData filter
3. Get list items with field selection
4. Create new list item
5. Update existing list item
6. Delete list item
7. Create item with invalid fields (error case)

### 9.2 Test Scenarios

```typescript
// Example test using Claude Desktop

// Scenario 1: Search for budget documents
User: "Search for documents with 'budget' in SharePoint"
Expected Tool Call: search_documents({ query: "budget" })
Expected Output: List of matching documents with metadata

// Scenario 2: Read file content
User: "Show me the content of file ID abc123"
Expected Tool Call: get_file_content({ file_id: "abc123" })
Expected Output: File content as text

// Scenario 3: Create list item
User: "Add a new task: 'Review Q1 budget' with status 'Not Started'"
Expected Tool Call: create_list_item({
  list_name: "Tasks",
  fields: { Title: "Review Q1 budget", Status: "Not Started" }
})
Expected Output: Created item ID and details
```

---

## 10. Implementation Checklist

### Phase 1: Core Infrastructure (Week 1)

- [ ] Initialize Cloudflare Workers project
- [ ] Set up TypeScript configuration
- [ ] Install MCP SDK dependencies
- [ ] Create KV namespace for token storage
- [ ] Implement `src/types/env.ts` (environment types)
- [ ] Implement `src/middleware/auth.ts` (API Key validation)
- [ ] Implement `src/auth.ts` (AuthManager class)
- [ ] Test Client Credentials flow manually
- [ ] Store secrets in Cloudflare Workers

### Phase 2: Graph API Client (Week 2)

- [ ] Implement `src/graph-client.ts` (GraphClient class)
- [ ] Implement endpoint mapping constants
- [ ] Test GET, POST, PATCH, DELETE methods
- [ ] Implement error handling and retry logic
- [ ] Test token caching in KV

### Phase 3: MCP Server & Tools (Week 2)

- [ ] Implement `src/server.ts` (createMCPServer)
- [ ] Implement `src/index.ts` (Cloudflare Worker entry)
- [ ] Register all 7 tools in server
- [ ] Implement `src/tools/documents.ts`
  - [ ] `handleSearchDocuments`
  - [ ] `handleGetFileContent`
  - [ ] `handleListFiles`
- [ ] Implement `src/tools/lists.ts`
  - [ ] `handleGetListItems`
  - [ ] `handleCreateListItem`
  - [ ] `handleUpdateListItem`
  - [ ] `handleDeleteListItem`

### Phase 4: Testing & Documentation (Week 3)

- [ ] Deploy to Cloudflare Workers
- [ ] Test with Claude Desktop MCP client
- [ ] Test all 7 tools with real SharePoint data
- [ ] Verify error handling for edge cases
- [ ] Monitor Cloudflare Workers analytics
- [ ] Write README with setup instructions
- [ ] Document Azure AD setup process
- [ ] Create usage examples

---

## 11. Next Steps

After design approval:
1. Run `/pdca do sharepoint-mcp-server` to get implementation guide
2. Initialize Cloudflare Workers project with `npx wrangler init`
3. Set up Azure AD app registration with Application permissions
4. Begin Week 1 implementation tasks

---

## References

- **Plan Document**: `docs/01-plan/features/sharepoint-mcp-server.plan.md`
- **MCP SDK TypeScript**: https://github.com/modelcontextprotocol/typescript-sdk
- **Cloudflare Workers Docs**: https://developers.cloudflare.com/workers
- **Microsoft Graph API Reference**: https://learn.microsoft.com/en-us/graph/api/overview
- **Streamable HTTP Transport**: https://spec.modelcontextprotocol.io/specification/2024-11-05/server/transports/#streamable-http
