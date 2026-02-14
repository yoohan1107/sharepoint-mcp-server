// SharePoint List CRUD MCP tools

import { Env } from "../types/env";
import { GraphClient, getSiteIdFromUrl, GRAPH_ENDPOINTS } from "../graph-client";
import {
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
} from "../types/models";

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
    created_by: item.createdBy.user.displayName,
    modified_by: item.lastModifiedBy.user.displayName,
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
