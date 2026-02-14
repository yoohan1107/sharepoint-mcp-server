// Environment variables and Cloudflare Workers bindings

export interface Env {
  // Secrets (설정은 Cloudflare Secrets로 관리)
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
