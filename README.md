# SharePoint MCP Server

Remote Model Context Protocol (MCP) server for Microsoft SharePoint Online integration.

## Features

- 🔍 Search SharePoint documents by keyword
- 📄 Download and extract file content
- 📁 List files and folders
- 📋 SharePoint List CRUD operations (Get, Create, Update, Delete)
- 🔐 Dual-layer authentication (API Key + Azure AD Client Credentials)
- ⚡ Deployed on Cloudflare Workers (Edge computing)

## Architecture

```
AI Client → MCP Server (CF Workers) → Microsoft Graph API → SharePoint Online
```

## Prerequisites

- Azure AD application with Application permissions:
  - `Sites.Read.All`
  - `Sites.ReadWrite.All`
  - `Files.Read.All`
- Cloudflare account
- SharePoint Online site access

## Deployment

This project uses GitHub Actions for automatic deployment to Cloudflare Workers.

### Setup

1. **Azure AD Configuration**
   - Register app in Azure Portal
   - Get Client ID, Client Secret, Tenant ID
   - Grant Application permissions with Admin consent

2. **Cloudflare Configuration**
   - Get API Token: Cloudflare Dashboard → My Profile → API Tokens → Create Token
   - Get Account ID: Cloudflare Dashboard → Workers & Pages (bottom right)

3. **GitHub Secrets**

   Add these secrets to your GitHub repository (Settings → Secrets and variables → Actions):

   ```
   CLOUDFLARE_API_TOKEN=<your-cloudflare-api-token>
   CLOUDFLARE_ACCOUNT_ID=<your-cloudflare-account-id>
   ```

4. **Cloudflare Workers Secrets**

   After first deployment, set these via Cloudflare Dashboard:

   Workers & Pages → Your Worker → Settings → Variables → Add:

   ```
   MCP_API_KEY=<your-api-key>
   AZURE_CLIENT_ID=<azure-app-client-id>
   AZURE_CLIENT_SECRET=<azure-app-secret>
   AZURE_TENANT_ID=<azure-tenant-id>
   SHAREPOINT_SITE_URL=https://yourtenant.sharepoint.com/sites/YourSite
   ```

5. **KV Namespace**

   Workers & Pages → Your Worker → Settings → Bindings → Add:
   - Variable name: `GRAPH_TOKEN_CACHE`
   - Type: KV Namespace
   - Create new namespace or select existing

6. **Deploy**

   ```bash
   git add .
   git commit -m "Initial commit"
   git push origin main
   ```

   GitHub Actions will automatically deploy to Cloudflare Workers.

## MCP Tools

### Document Tools

1. **search_documents**: Search documents by keyword
2. **get_file_content**: Download file content as text
3. **list_files**: List files in a folder

### SharePoint List Tools

4. **get_list_items**: Query list items with OData filters
5. **create_list_item**: Create new list item
6. **update_list_item**: Update existing list item
7. **delete_list_item**: Delete list item

## Usage with Claude Desktop

Add to your Claude Desktop MCP settings:

```json
{
  "mcpServers": {
    "sharepoint": {
      "url": "https://sharepoint-mcp-server.workers.dev/mcp",
      "headers": {
        "Authorization": "Bearer <your-mcp-api-key>"
      }
    }
  }
}
```

Then you can ask Claude:
- "Search for budget documents in SharePoint"
- "Show me the content of file ID abc123"
- "List items in the Tasks list where Status is Active"

## Development

**Note**: Local development requires x64 architecture. On ARM64 (Apple Silicon, Windows ARM), use GitHub Actions for deployment.

## License

MIT

## References

- [MCP Specification](https://spec.modelcontextprotocol.io)
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview)
- [Cloudflare Workers](https://developers.cloudflare.com/workers)
