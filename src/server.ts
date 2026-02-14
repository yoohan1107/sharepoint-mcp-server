// MCP Server instance with tool registration

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import {
  ListToolsRequestSchema,
  CallToolRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { Env } from "./types/env";

// Import tool handlers
import {
  handleGetSiteInfo,
  handleSearchDocuments,
  handleGetFileContent,
  handleListFiles,
} from "./tools/documents";
import {
  handleGetListItems,
  handleCreateListItem,
  handleUpdateListItem,
  handleDeleteListItem,
} from "./tools/lists";

/**
 * Create and configure MCP server with all SharePoint tools
 */
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
        // ============================================================
        // Site Info Tool
        // ============================================================
        {
          name: "get_site_info",
          description:
            "Get SharePoint site information including site name, description, and URL",
          inputSchema: {
            type: "object",
            properties: {
              site_url: {
                type: "string",
                description: "SharePoint site URL (defaults to configured site)",
              },
            },
            required: [],
          },
        },

        // ============================================================
        // Document Tools
        // ============================================================
        {
          name: "search_documents",
          description:
            "Search SharePoint documents by keyword with optional file type filter",
          inputSchema: {
            type: "object",
            properties: {
              query: {
                type: "string",
                description: "Search keyword or phrase",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
              file_type: {
                type: "string",
                description: "Optional file extension filter (e.g., 'docx', 'xlsx', 'pdf')",
              },
              max_results: {
                type: "number",
                description: "Maximum number of results to return (default 10, max 50)",
              },
            },
            required: ["query"],
          },
        },
        {
          name: "get_file_content",
          description:
            "Download and extract text content from a SharePoint file",
          inputSchema: {
            type: "object",
            properties: {
              file_id: {
                type: "string",
                description: "DriveItem ID of the file to download",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
            },
            required: ["file_id"],
          },
        },
        {
          name: "list_files",
          description:
            "List files and folders in a SharePoint directory",
          inputSchema: {
            type: "object",
            properties: {
              folder_path: {
                type: "string",
                description: "Folder path to list (default: root '/')",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
            },
            required: [],
          },
        },

        // ============================================================
        // SharePoint List Tools
        // ============================================================
        {
          name: "get_list_items",
          description:
            "Query SharePoint list items with OData filtering and field selection",
          inputSchema: {
            type: "object",
            properties: {
              list_name: {
                type: "string",
                description: "List name or ID",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
              filter: {
                type: "string",
                description: "OData filter query (e.g., \"Status eq 'Active'\")",
              },
              select: {
                type: "array",
                items: { type: "string" },
                description: "Array of field names to return",
              },
              top: {
                type: "number",
                description: "Maximum number of items to return (default 50, max 100)",
              },
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
              list_name: {
                type: "string",
                description: "List name or ID",
              },
              fields: {
                type: "object",
                description: "Field values to set (key-value pairs)",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
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
              list_name: {
                type: "string",
                description: "List name or ID",
              },
              item_id: {
                type: "string",
                description: "Item ID to update",
              },
              fields: {
                type: "object",
                description: "Field values to update (key-value pairs)",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
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
              list_name: {
                type: "string",
                description: "List name or ID",
              },
              item_id: {
                type: "string",
                description: "Item ID to delete",
              },
              site_id: {
                type: "string",
                description: "Optional SharePoint site ID (defaults to configured site)",
              },
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
      switch (name) {
        // Site info
        case "get_site_info":
          return await handleGetSiteInfo(env, args as any);

        // Document tools
        case "search_documents":
          return await handleSearchDocuments(env, args as any);

        case "get_file_content":
          return await handleGetFileContent(env, args as any);

        case "list_files":
          return await handleListFiles(env, args as any);

        // List tools
        case "get_list_items":
          return await handleGetListItems(env, args as any);

        case "create_list_item":
          return await handleCreateListItem(env, args as any);

        case "update_list_item":
          return await handleUpdateListItem(env, args as any);

        case "delete_list_item":
          return await handleDeleteListItem(env, args as any);

        default:
          throw new Error(`Unknown tool: ${name}`);
      }
    } catch (error) {
      console.error(`Tool execution error (${name}):`, error);

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
