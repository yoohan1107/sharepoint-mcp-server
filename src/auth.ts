// Azure AD Client Credentials authentication and token management

import { Env, GraphTokenCache } from "./types/env";

export class AuthManager {
  private env: Env;
  private tokenCacheKey = "graph_access_token";

  constructor(env: Env) {
    this.env = env;
  }

  /**
   * Get valid access token (from cache or request new)
   */
  async getAccessToken(): Promise<string> {
    // Check KV cache first
    const cached = await this.getCachedToken();

    if (cached && !this.isTokenExpired(cached)) {
      console.log("Using cached Graph API token");
      return cached.access_token;
    }

    console.log("Requesting new Graph API token");

    // Request new token
    const token = await this.requestNewToken();

    // Cache in KV
    await this.cacheToken(token);

    return token.access_token;
  }

  /**
   * Request new access token from Azure AD using Client Credentials Grant
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
      const errorText = await response.text();
      throw new Error(
        `Token request failed (${response.status}): ${errorText}`
      );
    }

    const data: any = await response.json();

    // Create token cache with 5-minute buffer before expiration
    const expiresIn = data.expires_in || 3600; // Default 1 hour
    const bufferSeconds = 300; // 5 minutes

    return {
      access_token: data.access_token,
      expires_at: Date.now() + (expiresIn - bufferSeconds) * 1000,
      token_type: "Bearer",
    };
  }

  /**
   * Get token from KV cache
   */
  private async getCachedToken(): Promise<GraphTokenCache | null> {
    try {
      const cached = await this.env.GRAPH_TOKEN_CACHE.get(
        this.tokenCacheKey,
        "json"
      );

      return cached as GraphTokenCache | null;
    } catch (error) {
      console.error("KV cache read error:", error);
      return null;
    }
  }

  /**
   * Store token in KV cache with TTL
   */
  private async cacheToken(token: GraphTokenCache): Promise<void> {
    try {
      const ttl = Math.floor((token.expires_at - Date.now()) / 1000);

      await this.env.GRAPH_TOKEN_CACHE.put(
        this.tokenCacheKey,
        JSON.stringify(token),
        { expirationTtl: ttl }
      );

      console.log(`Token cached with TTL: ${ttl} seconds`);
    } catch (error) {
      console.error("KV cache write error:", error);
      // Don't throw - continue with uncached token
    }
  }

  /**
   * Check if token is expired
   */
  private isTokenExpired(token: GraphTokenCache): boolean {
    return Date.now() >= token.expires_at;
  }
}
