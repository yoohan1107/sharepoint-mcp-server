# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SharePoint MCP Server is a Remote Model Context Protocol (MCP) server that enables AI clients (Claude, ChatGPT, Cursor, etc.) to directly access documents, files, and list data stored in Microsoft SharePoint Online. This is a PoC/personal learning project deployed on Cloudflare Workers (free tier).

**Key Characteristics:**
- **Runtime**: Cloudflare Workers with TypeScript
- **Protocol**: MCP (Model Context Protocol) via Streamable HTTP
- **Integration**: Microsoft Graph API v1.0 for SharePoint access
- **MCP Authentication**: API Key (header-based, simple validation)
- **Graph API Authentication**: Azure AD Client Credentials Grant (Application permissions)
- **Token Storage**: Cloudflare KV for Graph API token caching

## Technical Architecture

### Project Structure (Planned)

```
sharepoint-mcp-server/
├── src/
│   ├── index.ts              # Entry point, HTTP routing, API Key validation
│   ├── server.ts             # MCP server instance & Tool registration
│   ├── auth.ts               # Client Credentials flow, Graph API token management
│   ├── graph-client.ts       # Microsoft Graph API client
│   └── tools/
│       ├── documents.ts      # Document search/retrieval/download Tools
│       └── lists.ts          # SharePoint List CRUD Tools
├── wrangler.jsonc            # Cloudflare Workers configuration
├── package.json
└── tsconfig.json
```

### MCP Tools (Implemented)

**Site Tool:**
- `get_site_info` - Get SharePoint site information (name, description, URL)

**Document/File Tools:**
- `search_documents` - Search SharePoint documents by keyword
- `get_file_content` - Download file content and extract text
- `list_files` - List files in a specific folder

**SharePoint List Tools:**
- `get_list_items` - Query list items with OData filtering
- `create_list_item` - Create new list items
- `update_list_item` - Update existing list items
- `delete_list_item` - Delete list items

All tools should follow the MCP SDK tool specification pattern with proper parameter validation and error handling.

### Authentication Flow

**Two-Layer Authentication**:

**Layer 1: MCP Client → MCP Server (API Key)**
1. AI client sends request with API Key in Authorization header
2. MCP server validates API Key against `MCP_API_KEY` environment variable
3. If valid, proceed to Layer 2; if invalid, return 401 Unauthorized

**Layer 2: MCP Server → Graph API (Client Credentials)**
1. MCP server requests access token from Azure AD using Client Credentials Grant
2. Azure AD issues access token (valid for 1 hour)
3. MCP server caches token in Cloudflare KV
4. MCP server calls Microsoft Graph API with Bearer token
5. Graph API returns SharePoint data

**Required Azure AD Application Permissions** (admin consent required, no user interaction):
- `Sites.Read.All` - Read sites and documents
- `Sites.ReadWrite.All` - CRUD operations on lists
- `Files.Read.All` - Read file content

**Note**: Application permissions work without browser-based user authentication, suitable for server-to-server scenarios.

### Environment Variables (Secrets)

Configure these in Cloudflare Workers:
- `MCP_API_KEY` - API key for MCP client authentication
- `AZURE_CLIENT_ID` - Azure AD app client ID (with Application permissions)
- `AZURE_CLIENT_SECRET` - Azure AD app client secret
- `AZURE_TENANT_ID` - Azure AD tenant ID
- `SHAREPOINT_SITE_URL` - Default SharePoint site URL

## Development Commands

### Cloudflare Workers (Wrangler CLI)

```bash
# Install dependencies
npm install

# Local development (with bindings)
npx wrangler dev

# Deploy to production (수동 - 보통 불필요)
npx wrangler deploy

# Tail logs
npx wrangler tail

# Manage secrets
npx wrangler secret put MCP_API_KEY
npx wrangler secret put AZURE_CLIENT_ID
npx wrangler secret put AZURE_CLIENT_SECRET
npx wrangler secret put AZURE_TENANT_ID
npx wrangler secret put SHAREPOINT_SITE_URL

# Create KV namespace
npx wrangler kv:namespace create "AUTH_TOKENS"
```

### CI/CD (자동 배포)

**GitHub Actions로 자동 배포가 설정되어 있다** (`.github/workflows/deploy.yml`):
- `main` 브랜치에 push하면 자동으로 Cloudflare Workers에 배포됨
- 수동 `npx wrangler deploy`는 불필요 — **git commit & push만 하면 된다**
- 필요 시크릿: `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID` (GitHub Secrets에 설정됨)
- `workflow_dispatch`로 수동 트리거도 가능

### Testing

Since this is a PoC, testing will be manual:
1. Connect Claude Desktop or other MCP client to the deployed Worker URL
2. Test natural language commands that trigger MCP tools
3. Verify OAuth flow completes successfully
4. Validate SharePoint data is correctly retrieved/modified

## Microsoft Graph API Reference

Key endpoints used in this project:

| Operation | Method | Endpoint |
|-----------|--------|----------|
| Search documents | GET | `/sites/{site-id}/drive/root/search(q='{query}')` |
| List files | GET | `/sites/{site-id}/drive/root:/{path}:/children` |
| Download file | GET | `/sites/{site-id}/drive/items/{item-id}/content` |
| Get list items | GET | `/sites/{site-id}/lists/{list-id}/items` |
| Create list item | POST | `/sites/{site-id}/lists/{list-id}/items` |
| Update list item | PATCH | `/sites/{site-id}/lists/{list-id}/items/{item-id}` |
| Delete list item | DELETE | `/sites/{site-id}/lists/{list-id}/items/{item-id}` |

Use `@microsoft/microsoft-graph-client` or direct `fetch()` calls with Bearer token.

## MCP SDK Integration

This project uses `@modelcontextprotocol/sdk` for TypeScript:

```typescript
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamable.js";

// Register tools with proper schema
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: [...] };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  // Tool execution logic
});
```

Transport must be Streamable HTTP (not SSE) for Cloudflare Workers compatibility.

## Important Constraints

### Cloudflare Workers Free Tier Limits
- 100,000 requests/day
- 10ms CPU time per request
- 128MB memory
- No long-running processes (max 30s wall time on free tier)

### Rate Limiting
Microsoft Graph API has throttling limits. Implement:
- Token caching in KV to minimize auth requests
- Request retry with exponential backoff
- Batch requests where possible

### Security Considerations
- Never log or expose OAuth tokens
- Use HTTPS only for OAuth callbacks
- Set appropriate token TTL in KV storage
- Validate all user inputs to prevent injection attacks

## PDCA Workflow

This project uses the bkit PDCA (Plan-Do-Check-Act) methodology. Current status is tracked in `docs/.pdca-status.json`:

- **Current Phase**: 1 (Schema/Terminology)
- **Level**: Dynamic (fullstack with BaaS patterns)

When implementing features, follow the 9-phase development pipeline:
1. Schema - Define terminology and data structures
2. Convention - Establish coding standards
3. Mockup - Create UI prototypes (if applicable)
4. API - Implement backend Tools
5. Design System - Component patterns
6. UI Integration - Connect frontend to Tools
7. SEO/Security - Security hardening and optimization
8. Review - Gap analysis between design and implementation
9. Deployment - Production deployment to Cloudflare Workers

## References

- [MCP Specification](https://spec.modelcontextprotocol.io)
- [MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/api/overview)
- [Cloudflare Workers Documentation](https://developers.cloudflare.com/workers)
- [Cloudflare MCP Template](https://github.com/cloudflare/ai/tree/main/demos/remote-mcp-server)
- Product Requirements: `PRD/Initial_prd.md`
