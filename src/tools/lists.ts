// SharePoint List CRUD MCP tools

import { Env } from "../types/env";
import { GraphClient, getSiteIdFromUrl, GRAPH_ENDPOINTS } from "../graph-client";
import {
  ListListsInput,
  ListListsOutput,
  GetListColumnsInput,
  GetListColumnsOutput,
  GetListItemInput,
  GetListItemOutput,
  GetListItemsInput,
  GetListItemsOutput,
  CreateListItemInput,
  CreateListItemOutput,
  UpdateListItemInput,
  UpdateListItemOutput,
  DeleteListItemInput,
  DeleteListItemOutput,
  ListItem,
  GraphListItem,
  GraphListResponse,
  GraphSiteListResponse,
  GraphColumnResponse,
  GraphColumnDefinition,
} from "../types/models";

const COLUMN_TYPE_KEYS = [
  "text",
  "number",
  "choice",
  "dateTime",
  "boolean",
  "currency",
  "lookup",
  "personOrGroup",
  "hyperlinkOrPicture",
  "calculated",
] as const;

function getDisplayName(actor?: { user?: { displayName?: string } }): string {
  return actor?.user?.displayName || "Unknown";
}

function getColumnType(column: GraphColumnDefinition): string {
  for (const key of COLUMN_TYPE_KEYS) {
    if (column[key]) {
      return key;
    }
  }

  return "unknown";
}

/**
 * Tool: list_lists
 * List available SharePoint lists and libraries in a site
 */
export async function handleListLists(
  env: Env,
  args: ListListsInput
): Promise<any> {
  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));
  const top = Math.min(args.top || 100, 200);
  const includeHidden = args.include_hidden || false;

  const params = new URLSearchParams();
  params.set(
    "$select",
    "id,name,displayName,description,webUrl,createdDateTime,lastModifiedDateTime,list"
  );
  params.set("$top", String(top));

  const endpoint = `${GRAPH_ENDPOINTS.LIST_LISTS(siteId)}?${params}`;
  const response = await client.get<GraphSiteListResponse>(endpoint);

  const lists = response.value
    .map((list) => ({
      id: list.id,
      name: list.name,
      display_name: list.displayName,
      description: list.description || "",
      web_url: list.webUrl,
      template: list.list?.template || "unknown",
      is_hidden: list.list?.hidden || false,
      created_datetime: list.createdDateTime,
      last_modified_datetime: list.lastModifiedDateTime,
    }))
    .filter((list) => includeHidden || !list.is_hidden);

  const result: ListListsOutput = {
    success: true,
    lists,
    count: lists.length,
    has_more: !!response["@odata.nextLink"],
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: get_list_columns
 * List columns (schema) for a SharePoint list
 */
export async function handleGetListColumns(
  env: Env,
  args: GetListColumnsInput
): Promise<any> {
  if (!args.list_name) {
    throw new Error("list_name is required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));
  const top = Math.min(args.top || 200, 500);

  const params = new URLSearchParams();
  params.set(
    "$select",
    "id,name,displayName,description,required,readOnly,hidden,columnGroup,text,number,choice,dateTime,boolean,currency,lookup,personOrGroup,hyperlinkOrPicture,calculated"
  );
  params.set("$top", String(top));

  const endpoint = `${GRAPH_ENDPOINTS.GET_LIST_COLUMNS(siteId, args.list_name)}?${params}`;
  const response = await client.get<GraphColumnResponse>(endpoint);

  const columns = response.value.map((column) => ({
    id: column.id,
    name: column.name,
    display_name: column.displayName,
    description: column.description || "",
    type: getColumnType(column),
    required: !!column.required,
    read_only: !!column.readOnly,
    hidden: !!column.hidden,
    column_group: column.columnGroup || "",
  }));

  const result: GetListColumnsOutput = {
    success: true,
    list_name: args.list_name,
    columns,
    count: columns.length,
    has_more: !!response["@odata.nextLink"],
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: get_list_item
 * Get a single SharePoint list item by item ID
 */
export async function handleGetListItem(
  env: Env,
  args: GetListItemInput
): Promise<any> {
  if (!args.list_name || !args.item_id) {
    throw new Error("list_name and item_id are required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));

  const params = new URLSearchParams();
  params.set("$expand", "fields");

  const endpoint = `${GRAPH_ENDPOINTS.GET_LIST_ITEM(siteId, args.list_name, args.item_id)}?${params}`;
  const response = await client.get<GraphListItem>(endpoint);

  const item: ListItem = {
    id: response.id,
    fields: response.fields,
    created_datetime: response.createdDateTime,
    last_modified_datetime: response.lastModifiedDateTime,
    created_by: getDisplayName(response.createdBy),
    modified_by: getDisplayName(response.lastModifiedBy),
  };

  const result: GetListItemOutput = {
    success: true,
    list_name: args.list_name,
    item,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: get_list_items
 * Query SharePoint list items with OData filtering
 */
export async function handleGetListItems(
  env: Env,
  args: GetListItemsInput
): Promise<any> {
  if (!args.list_name) {
    throw new Error("list_name is required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));
  const top = Math.min(args.top || 50, 100);

  // Build query parameters
  const params = new URLSearchParams();
  params.set("$expand", "fields");
  params.set("$top", String(top));

  if (args.filter) {
    params.set("$filter", args.filter);
  }

  if (args.select && args.select.length > 0) {
    params.set("$select", args.select.join(","));
  }

  const endpoint = `${GRAPH_ENDPOINTS.GET_LIST_ITEMS(siteId, args.list_name)}?${params}`;
  const response = await client.get<GraphListResponse>(endpoint);

  const items: ListItem[] = response.value.map((item) => ({
    id: item.id,
    fields: item.fields,
    created_datetime: item.createdDateTime,
    last_modified_datetime: item.lastModifiedDateTime,
    created_by: getDisplayName(item.createdBy),
    modified_by: getDisplayName(item.lastModifiedBy),
  }));

  const result: GetListItemsOutput = {
    success: true,
    list_name: args.list_name,
    items,
    count: items.length,
    has_more: !!response["@odata.nextLink"],
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: create_list_item
 * Create a new item in a SharePoint list
 */
export async function handleCreateListItem(
  env: Env,
  args: CreateListItemInput
): Promise<any> {
  if (!args.list_name || !args.fields) {
    throw new Error("list_name and fields are required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));

  const endpoint = GRAPH_ENDPOINTS.CREATE_LIST_ITEM(siteId, args.list_name);
  const body = { fields: args.fields };

  const response = await client.post<GraphListItem>(endpoint, body);

  const result: CreateListItemOutput = {
    success: true,
    item_id: response.id,
    fields: response.fields,
    created_datetime: response.createdDateTime,
    web_url: `${env.SHAREPOINT_SITE_URL}/Lists/${args.list_name}/DispForm.aspx?ID=${response.id}`,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: update_list_item
 * Update an existing SharePoint list item
 */
export async function handleUpdateListItem(
  env: Env,
  args: UpdateListItemInput
): Promise<any> {
  if (!args.list_name || !args.item_id || !args.fields) {
    throw new Error("list_name, item_id, and fields are required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));

  const endpoint = GRAPH_ENDPOINTS.UPDATE_LIST_ITEM(siteId, args.list_name, args.item_id);
  const response = await client.patch<Record<string, any>>(endpoint, args.fields);

  const result: UpdateListItemOutput = {
    success: true,
    item_id: args.item_id,
    fields: response,
    last_modified_datetime: new Date().toISOString(),
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}

/**
 * Tool: delete_list_item
 * Delete a SharePoint list item
 */
export async function handleDeleteListItem(
  env: Env,
  args: DeleteListItemInput
): Promise<any> {
  if (!args.list_name || !args.item_id) {
    throw new Error("list_name and item_id are required");
  }

  const client = new GraphClient(env);
  const siteId = args.site_id || (await getSiteIdFromUrl(client, env.SHAREPOINT_SITE_URL));

  const endpoint = GRAPH_ENDPOINTS.DELETE_LIST_ITEM(siteId, args.list_name, args.item_id);
  await client.delete(endpoint);

  const result: DeleteListItemOutput = {
    success: true,
    item_id: args.item_id,
    deleted_at: new Date().toISOString(),
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
}
