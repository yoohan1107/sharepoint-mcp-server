// Cloudflare Worker entry point for SharePoint MCP Server

import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamable.js";
import { createMCPServer } from "./server";
import { validateAPIKey } from "./middleware/auth";
import { Env } from "./types/env";

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    // CORS headers for browser clients
    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
    };

    // Handle preflight requests
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

    // Validate MCP API Key (Layer 1 Authentication)
    const isValid = await validateAPIKey(request, env);

    if (!isValid) {
      console.error("API Key validation failed");

      return new Response(
        JSON.stringify({
          error: "Invalid or missing API key",
          message: "Authorization header must contain valid Bearer token",
        }),
        {
          status: 401,
          headers: {
            "Content-Type": "application/json",
            ...corsHeaders,
          },
        }
      );
    }

    console.log("API Key validated successfully");

    try {
      // Create MCP server instance
      const server = createMCPServer(env);

      // Create Streamable HTTP transport
      const transport = new StreamableHTTPServerTransport({
        endpoint: "/mcp",
      });

      // Connect server to transport
      await server.connect(transport);

      // Handle MCP request
      const mcpResponse = await transport.handleRequest(request);

      // Add CORS headers to response
      const headers = new Headers(mcpResponse.headers);
      Object.entries(corsHeaders).forEach(([key, value]) => {
        headers.set(key, value);
      });

      return new Response(mcpResponse.body, {
        status: mcpResponse.status,
        statusText: mcpResponse.statusText,
        headers,
      });
    } catch (error) {
      console.error("MCP Server error:", error);

      return new Response(
        JSON.stringify({
          error: "Internal server error",
          message: error instanceof Error ? error.message : String(error),
        }),
        {
          status: 500,
          headers: {
            "Content-Type": "application/json",
            ...corsHeaders,
          },
        }
      );
    }
  },
};
