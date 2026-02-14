// Document-related MCP tools: search, get content, list files

import { Env } from "../types/env";
import { GraphClient, getSiteIdFromUrl, GRAPH_ENDPOINTS } from "../graph-client";
import {
  SearchDocumentsInput,
  SearchDocumentsOutput,
  GetFileContentInput,
  GetFileContentOutput,
  ListFilesInput,
  ListFilesOutput,
  DocumentMetadata,
  FileSystemItem,
  GraphDriveItem,
  GraphSearchResponse,
} from "../types/models";

/**
 * Tool: search_documents
 * Search SharePoint documents by keyword with optional filters
 */
export async function handleSearchDocuments(
  env: Env,
  args: SearchDocumentsInput
): Promise<any> {
  // Validate input
  if (!args.query || args.query.trim() === "") {
    throw new Error("Query parameter is required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));
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
  const documents: DocumentMetadata[] = response.value.map((item) => ({
    id: item.id,
    name: item.name,
    web_url: item.webUrl,
    created_datetime: item.createdDateTime,
    last_modified_datetime: item.lastModifiedDateTime,
    size: item.size,
    file_type: item.name.split(".").pop() || "",
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
 * Tool: get_file_content
 * Download and extract text content from a SharePoint file
 */
export async function handleGetFileContent(
  env: Env,
  args: GetFileContentInput
): Promise<any> {
  if (!args.file_id) {
    throw new Error("file_id is required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));

  // Get file metadata first
  const metadata = await client.get<GraphDriveItem>(
    `/sites/${siteId}/drive/items/${args.file_id}`
  );

  // Download file content
  const contentEndpoint = GRAPH_ENDPOINTS.GET_FILE_CONTENT(siteId, args.file_id);
  const authManager = new (await import("../auth")).AuthManager(env);
  const token = await authManager.getAccessToken();

  const contentResponse = await fetch(
    `https://graph.microsoft.com/v1.0${contentEndpoint}`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
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
 * Tool: list_files
 * List files and folders in a SharePoint directory
 */
export async function handleListFiles(
  env: Env,
  args: ListFilesInput
): Promise<any> {
  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));
  const folderPath = args.folder_path || "/";

  const endpoint = GRAPH_ENDPOINTS.LIST_FILES(siteId, folderPath);
  const response = await client.get<GraphSearchResponse>(endpoint);

  const items: FileSystemItem[] = response.value.map((item) => ({
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
