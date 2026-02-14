// Microsoft Graph API client with authentication and error handling

import { Env } from "./types/env";
import { AuthManager } from "./auth";
import { ErrorCode, MCPError } from "./types/models";

export class GraphClient {
  private baseUrl = "https://graph.microsoft.com/v1.0";
  private authManager: AuthManager;

  constructor(env: Env) {
    this.authManager = new AuthManager(env);
  }

  /**
   * Generic GET request to Graph API
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
   * Generic POST request to Graph API
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
   * Generic PATCH request to Graph API
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
   * Generic DELETE request to Graph API
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
   * Handle error response and create MCPError
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
      message:
        errorData.error?.message || errorData.message || "Graph API error",
      details: errorData,
    };

    return new Error(JSON.stringify(error));
  }

  /**
   * Map HTTP status code to error code
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

// ============================================================================
// Graph API Endpoint Helpers
// ============================================================================

/**
 * Get site ID from SharePoint URL
 */
export async function getSiteIdFromUrl(
  client: GraphClient,
  siteUrl: string
): Promise<string> {
  const url = new URL(siteUrl);
  const endpoint = `/sites/${url.hostname}:${url.pathname}`;

  const site = await client.get<{ id: string }>(endpoint);
  return site.id;
}

/**
 * Graph API endpoint templates
 */
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
  GET_SITE_BY_URL: (siteUrl: string) => {
    const url = new URL(siteUrl);
    return `/sites/${url.hostname}:${url.pathname}`;
  },
};
