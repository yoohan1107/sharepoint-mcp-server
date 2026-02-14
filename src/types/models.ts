// Data models for MCP tools and Microsoft Graph API

// ============================================================================
// MCP Tool Input/Output Schemas
// ============================================================================

// Document Tools
export interface SearchDocumentsInput {
  query: string;
  site_id?: string;
  file_type?: string;
  max_results?: number;
}

export interface SearchDocumentsOutput {
  success: boolean;
  documents: DocumentMetadata[];
  total_count: number;
  query: string;
}

export interface GetFileContentInput {
  file_id: string;
  site_id?: string;
}

export interface GetFileContentOutput {
  success: boolean;
  file_id: string;
  file_name: string;
  content_type: string;
  content: string;
  size: number;
  encoding: string;
}

export interface ListFilesInput {
  folder_path?: string;
  site_id?: string;
}

export interface ListFilesOutput {
  success: boolean;
  folder_path: string;
  items: FileSystemItem[];
  count: number;
}

// Site Info Tool
export interface GetSiteInfoInput {
  site_url?: string;
}

export interface GetSiteInfoOutput {
  success: boolean;
  id: string;
  name: string;
  display_name: string;
  description: string;
  web_url: string;
  created_datetime: string;
  last_modified_datetime: string;
}

// SharePoint List Tools
export interface GetListItemsInput {
  list_name: string;
  site_id?: string;
  filter?: string;
  select?: string[];
  top?: number;
}

export interface GetListItemsOutput {
  success: boolean;
  list_name: string;
  items: ListItem[];
  count: number;
  has_more: boolean;
}

export interface CreateListItemInput {
  list_name: string;
  fields: Record<string, any>;
  site_id?: string;
}

export interface CreateListItemOutput {
  success: boolean;
  item_id: string;
  fields: Record<string, any>;
  created_datetime: string;
  web_url: string;
}

export interface UpdateListItemInput {
  list_name: string;
  item_id: string;
  fields: Record<string, any>;
  site_id?: string;
}

export interface UpdateListItemOutput {
  success: boolean;
  item_id: string;
  fields: Record<string, any>;
  last_modified_datetime: string;
}

export interface DeleteListItemInput {
  list_name: string;
  item_id: string;
  site_id?: string;
}

export interface DeleteListItemOutput {
  success: boolean;
  item_id: string;
  deleted_at: string;
}

// ============================================================================
// Common Data Models
// ============================================================================

export interface DocumentMetadata {
  id: string;
  name: string;
  web_url: string;
  created_datetime: string;
  last_modified_datetime: string;
  size: number;
  file_type: string;
  created_by: string;
  modified_by: string;
  parent_path: string;
}

export interface FileSystemItem {
  id: string;
  name: string;
  type: "file" | "folder";
  size?: number;
  last_modified: string;
  web_url: string;
}

export interface ListItem {
  id: string;
  fields: Record<string, any>;
  created_datetime: string;
  last_modified_datetime: string;
  created_by: string;
  modified_by: string;
}

// ============================================================================
// Microsoft Graph API Response Types
// ============================================================================

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

export interface GraphSiteResponse {
  id: string;
  name: string;
  displayName: string;
  description: string;
  webUrl: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
}

// ============================================================================
// Error Models
// ============================================================================

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
