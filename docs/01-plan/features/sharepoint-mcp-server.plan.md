# Plan: SharePoint MCP Server

**Feature Name**: SharePoint MCP Server
**Level**: Dynamic
**Type**: Core Infrastructure
**Priority**: P0 (Must-Have)
**Created**: 2026-02-14
**Status**: Planning

---

## 1. Feature Overview

### 1.1 Description

Build a Remote Model Context Protocol (MCP) server that enables AI clients (Claude, ChatGPT, Cursor, etc.) to directly access and interact with Microsoft SharePoint Online data. This server will expose SharePoint documents, files, and list data through standardized MCP tools, eliminating the need for manual copy-paste workflows between SharePoint and AI tools.

### 1.2 Business Value

- **Productivity**: Users can query SharePoint data directly from AI conversations without context switching
- **Automation**: AI can read/write SharePoint list data, enabling automated workflows
- **Integration**: Single MCP server works with multiple AI platforms (Claude, ChatGPT, Cursor)
- **Learning**: PoC project to understand MCP protocol and Cloudflare Workers deployment

### 1.3 Success Criteria

| # | Criterion | Measurement |
|---|-----------|-------------|
| 1 | AI client can search SharePoint documents via natural language | Manual test in Claude Desktop |
| 2 | SharePoint List CRUD operations work through MCP tools | Manual test with sample list |
| 3 | OAuth 2.0 authentication flow completes successfully | Auth flow test |
| 4 | Server runs within Cloudflare Workers free tier limits | CF Dashboard monitoring |
| 5 | Works with both Claude Desktop and compatible MCP clients | Cross-client testing |

---

## 2. Technical Scope

### 2.1 In Scope (Phase 1 - PoC)

**Core Infrastructure**:
- ✅ Cloudflare Workers runtime setup with TypeScript
- ✅ MCP server instance using @modelcontextprotocol/sdk
- ✅ Streamable HTTP transport (MCP latest standard)
- ✅ Azure AD OAuth 2.0 authentication flow
- ✅ Token caching in Cloudflare KV
- ✅ Microsoft Graph API client integration

**MCP Tools - Documents/Files** (P0):
- ✅ `search_documents` - Keyword-based document search
- ✅ `get_file_content` - Download and extract file text content
- ✅ `list_files` - List files in a folder

**MCP Tools - SharePoint Lists**:
- ✅ `get_list_items` - Query list items with OData filtering (P0)
- ✅ `create_list_item` - Create new list items (P1)
- ✅ `update_list_item` - Update existing items (P1)
- ✅ `delete_list_item` - Delete list items (P2)

**Error Handling**:
- ✅ OAuth token expiration/refresh logic
- ✅ Graph API rate limiting with retry/backoff
- ✅ Input validation for all MCP tool parameters
- ✅ Proper error messages returned via MCP protocol

### 2.2 Out of Scope (Future Phases)

- ❌ Document upload/creation (write operations for files)
- ❌ SharePoint site/page management (CMS features)
- ❌ Permission/access control management
- ❌ SharePoint On-Premise support (cloud-only)
- ❌ Batch operations for large datasets
- ❌ MCP Resources and Prompts (Tool-only for PoC)
- ❌ Webhook integration for real-time notifications

---

## 3. Architecture Plan

### 3.1 System Components

```
┌──────────────────┐     ┌─────────────────────┐     ┌──────────────────┐     ┌──────────────────┐
│   AI Client      │────▶│  Cloudflare Workers  │────▶│  Microsoft Graph │────▶│   SharePoint     │
│ (Claude/ChatGPT) │◀────│   (MCP Server)       │◀────│     API v1.0     │◀────│    Online        │
└──────────────────┘     └─────────────────────┘     └──────────────────┘     └──────────────────┘
  MCP + API Key         Streamable HTTP         Client Credentials Flow      REST API
```

### 3.2 Technology Stack

| Layer | Technology | Rationale |
|-------|-----------|-----------|
| **Runtime** | Cloudflare Workers | Free tier, global edge deployment, no cold starts |
| **Language** | TypeScript | MCP SDK official support, type safety |
| **MCP SDK** | @modelcontextprotocol/sdk | Anthropic official SDK for MCP servers |
| **Transport** | Streamable HTTP | Latest MCP standard, replaces SSE |
| **API** | Microsoft Graph API v1.0 | Unified endpoint for SharePoint access |
| **MCP Auth** | API Key | Simple header-based authentication for MCP clients |
| **Graph API Auth** | Azure AD Client Credentials | Application-level permissions, no user login required |
| **Cache** | Cloudflare KV | Token storage, included in free tier |

### 3.3 File Structure

```
sharepoint-mcp-server/
├── src/
│   ├── index.ts              # Entry point, HTTP routing, API Key validation
│   ├── server.ts             # MCP server instance, tool registration
│   ├── auth.ts               # Client Credentials flow, Graph API token management
│   ├── graph-client.ts       # Microsoft Graph API wrapper
│   └── tools/
│       ├── documents.ts      # Document search/retrieval/download
│       └── lists.ts          # SharePoint List CRUD operations
├── wrangler.jsonc            # Cloudflare Workers config
├── package.json
├── tsconfig.json
└── README.md
```

### 3.4 Authentication Flow

**Two-Layer Authentication**:

```
Layer 1: MCP Client Authentication (API Key)
──────────────────────────────────────────────
1. AI Client → MCP Server: Request with API Key in header
   Headers: { "Authorization": "Bearer z-rvzgY14wh1h4HJxCLhyeztjQdh62mq7wjIrvEwWzCksAF_cfskExJc01t9NURx" }
2. MCP Server: Validate API Key
3. If valid → proceed to Layer 2
   If invalid → return 401 Unauthorized

Layer 2: SharePoint Authentication (Client Credentials)
──────────────────────────────────────────────
4. MCP Server → Azure AD: Request token (Client Credentials Grant)
   POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
   Body: { client_id, client_secret, scope: "https://graph.microsoft.com/.default" }
5. Azure AD → MCP Server: Access token issued (valid 1 hour)
6. MCP Server → Cloudflare KV: Cache token with TTL
7. MCP Server → Graph API: Call with Bearer token
8. Graph API → SharePoint: Data operation
9. SharePoint → Graph API → MCP Server → AI Client: Response
```

**Required Azure AD Application Permissions** (no user consent needed):
- `Sites.Read.All` - Read sites and documents
- `Sites.ReadWrite.All` - CRUD on SharePoint lists
- `Files.Read.All` - Read file content

**Note**: Application permissions require admin consent but work without user interaction.

---

## 4. Implementation Plan

### 4.1 Development Phases (3 Weeks)

#### Week 1: Environment Setup & Basic Infrastructure
- [ ] Azure AD app registration with Application permissions
- [ ] Cloudflare Workers project initialization
- [ ] Install dependencies: MCP SDK, TypeScript, Wrangler CLI
- [ ] Create KV namespace for Graph API token cache
- [ ] Configure secrets (API Key, Azure credentials, SharePoint site URL)
- [ ] Implement basic MCP server with Streamable HTTP transport
- [ ] Implement API Key validation middleware
- [ ] Implement Client Credentials flow for Graph API
- [ ] Test dual-layer authentication

**Deliverable**: Working API Key + Client Credentials authentication

#### Week 2: MCP Tool Implementation
- [ ] Implement Microsoft Graph API client wrapper
- [ ] Implement Document Tools:
  - [ ] `search_documents` - Search with keyword + filters
  - [ ] `get_file_content` - Download and extract text
  - [ ] `list_files` - Folder browsing
- [ ] Implement List Tools:
  - [ ] `get_list_items` - Query with OData filters
  - [ ] `create_list_item` - Create new items
  - [ ] `update_list_item` - Update existing items
  - [ ] `delete_list_item` - Delete items
- [ ] Add input validation for all tools
- [ ] Implement error handling (token refresh, rate limiting, API errors)

**Deliverable**: All 7 MCP tools functional and tested

#### Week 3: Integration Testing & Documentation
- [ ] Test with Claude Desktop MCP client
- [ ] Test with other MCP-compatible clients (if available)
- [ ] Test natural language commands (e.g., "search for budget documents")
- [ ] Verify SharePoint List CRUD workflow end-to-end
- [ ] Monitor Cloudflare Workers metrics (requests, CPU time)
- [ ] Write README with setup instructions
- [ ] Document Azure AD setup process
- [ ] Document environment variables and secrets
- [ ] Create usage examples for each tool

**Deliverable**: Fully tested PoC with documentation

### 4.2 Key Files to Create

| File | Purpose | Priority |
|------|---------|----------|
| `src/index.ts` | Cloudflare Worker entry point, API Key validation | P0 |
| `src/server.ts` | MCP server setup + tool registration | P0 |
| `src/auth.ts` | Client Credentials flow, Graph API token cache | P0 |
| `src/graph-client.ts` | Graph API wrapper with token handling | P0 |
| `src/tools/documents.ts` | Document-related MCP tools | P0 |
| `src/tools/lists.ts` | SharePoint List MCP tools | P1 |
| `wrangler.jsonc` | Cloudflare Workers configuration | P0 |
| `package.json` | Dependencies and scripts | P0 |
| `tsconfig.json` | TypeScript configuration | P0 |
| `README.md` | Setup and usage documentation | P1 |

---

## 5. Dependencies

### 5.1 External Services

| Service | Purpose | Account Required |
|---------|---------|------------------|
| **Azure AD** | Application-level authentication (Client Credentials) | Microsoft/Office 365 admin account |
| **SharePoint Online** | Data source | SharePoint site access |
| **Microsoft Graph API** | API access to SharePoint | Azure AD app registration |
| **Cloudflare Workers** | Server runtime | Cloudflare account (free tier) |
| **Cloudflare KV** | Token storage | Included in Workers setup |

### 5.2 NPM Dependencies

```json
{
  "dependencies": {
    "@modelcontextprotocol/sdk": "^latest",
    "@cloudflare/workers-types": "^latest"
  },
  "devDependencies": {
    "typescript": "^5.x",
    "wrangler": "^3.x"
  }
}
```

### 5.3 Environment Variables (Secrets)

| Variable | Description | Source |
|----------|-------------|--------|
| `MCP_API_KEY` | API key for MCP client authentication | Provided: `z-rvzgY14wh1h4HJxCLhyeztjQdh62mq7wjIrvEwWzCksAF_cfskExJc01t9NURx` |
| `AZURE_CLIENT_ID` | Azure AD app client ID (Application permissions) | Azure Portal → App Registration |
| `AZURE_CLIENT_SECRET` | Azure AD app client secret | Azure Portal → Certificates & secrets |
| `AZURE_TENANT_ID` | Azure AD tenant ID | Azure Portal → Overview |
| `SHAREPOINT_SITE_URL` | Default SharePoint site URL | SharePoint site settings |

**Wrangler Secret Setup**:
```bash
npx wrangler secret put MCP_API_KEY
npx wrangler secret put AZURE_CLIENT_ID
npx wrangler secret put AZURE_CLIENT_SECRET
npx wrangler secret put AZURE_TENANT_ID
npx wrangler secret put SHAREPOINT_SITE_URL
```

---

## 6. Constraints & Risks

### 6.1 Technical Constraints

**Cloudflare Workers Free Tier**:
- 100,000 requests/day limit
- 10ms CPU time per request
- 128MB memory limit
- 30s wall time limit (free tier)

**Microsoft Graph API**:
- Rate limiting: Throttling varies by endpoint
- Token expiration: 1 hour for access tokens
- Refresh token: 90 days expiration

### 6.2 Risks & Mitigation

| Risk | Impact | Probability | Mitigation |
|------|--------|-------------|------------|
| Azure AD permission denied | High | Medium | Document required permissions, provide admin consent guide |
| CF Workers free tier exceeded | Low | Low | Monitor usage, implement caching, PoC scope is minimal |
| Graph API rate limiting | Medium | Medium | Implement retry with backoff, cache frequent queries |
| MCP protocol changes | Low | Low | Pin MCP SDK version, monitor breaking changes |
| OAuth token security | High | Low | Use KV encryption, HTTPS only, proper token TTL |

---

## 7. Testing Strategy

### 7.1 Manual Testing Checklist

**Authentication**:
- [ ] OAuth flow initiates correctly
- [ ] User can login with Microsoft account
- [ ] Access token is cached in KV
- [ ] Token refresh works on expiration

**Document Tools**:
- [ ] Search returns relevant documents
- [ ] File content extraction works for txt/csv/json
- [ ] File listing shows correct metadata

**List Tools**:
- [ ] Get list items with filters
- [ ] Create new list item
- [ ] Update existing list item
- [ ] Delete list item

**AI Integration**:
- [ ] Claude Desktop can connect to MCP server
- [ ] Natural language queries trigger correct tools
- [ ] Error messages are user-friendly

### 7.2 Edge Cases to Test

- Empty search results
- Invalid file IDs
- Non-existent SharePoint lists
- Expired/invalid OAuth tokens
- Malformed user inputs
- Rate limit exceeded scenario

---

## 8. Success Metrics

| Metric | Target | Measurement Method |
|--------|--------|-------------------|
| **Authentication Success Rate** | 100% | Manual auth flow tests |
| **Tool Success Rate** | >95% | Tool call logs during testing |
| **Response Time** | <2s per tool call | Cloudflare Workers analytics |
| **API Error Rate** | <5% | Graph API response logs |
| **Free Tier Compliance** | Within limits | CF Dashboard monitoring |

---

## 9. Future Enhancements (Post-PoC)

### Phase 2 Features
- Document upload capability (write operations)
- SharePoint site/page management tools
- MCP Resources for browsing SharePoint structure
- MCP Prompts for common SharePoint queries

### Phase 3 Features
- Webhook support for real-time notifications
- Batch operations for large data sets
- SharePoint On-Premise support (hybrid)
- Multi-tenant support for organization-wide deployment

### Alternative Deployment
- Migrate to Azure Functions (if internal infrastructure required)
- Add monitoring/logging with Azure Application Insights

---

## 10. References

- **PRD**: `PRD/Initial_prd.md`
- **MCP Specification**: https://spec.modelcontextprotocol.io
- **MCP TypeScript SDK**: https://github.com/modelcontextprotocol/typescript-sdk
- **Microsoft Graph API**: https://learn.microsoft.com/en-us/graph/api/overview
- **Cloudflare Workers**: https://developers.cloudflare.com/workers
- **Cloudflare MCP Template**: https://github.com/cloudflare/ai/tree/main/demos/remote-mcp-server

---

## Next Steps

After plan approval:
1. Run `/pdca design sharepoint-mcp-server` to create detailed design document
2. Set up Azure AD app registration
3. Initialize Cloudflare Workers project
4. Begin Week 1 implementation tasks
