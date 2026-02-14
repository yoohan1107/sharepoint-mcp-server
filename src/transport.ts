// Custom MCP Transport for Cloudflare Workers (Web API compatible)

import { Transport } from "@modelcontextprotocol/sdk/shared/transport.js";
import { JSONRPCMessage } from "@modelcontextprotocol/sdk/types.js";

/**
 * Check if a JSON-RPC message is a request (has "id" field) vs a notification (no "id").
 * Requests expect a response; notifications do not.
 */
function isRequest(message: JSONRPCMessage): boolean {
  return "id" in message && (message as any).id !== undefined;
}

export class WorkerTransport implements Transport {
  private responseResolve: ((value: JSONRPCMessage) => void) | null = null;

  onclose?: () => void;
  onerror?: (error: Error) => void;
  onmessage?: (message: JSONRPCMessage) => void;

  async start(): Promise<void> {}

  async close(): Promise<void> {
    this.onclose?.();
  }

  async send(message: JSONRPCMessage): Promise<void> {
    if (this.responseResolve) {
      this.responseResolve(message);
      this.responseResolve = null;
    }
  }

  async handleRequest(request: Request): Promise<Response> {
    const body = (await request.json()) as JSONRPCMessage;

    // Notifications (no "id") don't expect a response
    if (!isRequest(body)) {
      this.onmessage?.(body);
      return new Response(null, { status: 202 });
    }

    // Requests (with "id") expect a JSON-RPC response
    const response = await new Promise<JSONRPCMessage>((resolve) => {
      this.responseResolve = resolve;
      this.onmessage?.(body);
    });

    return new Response(JSON.stringify(response), {
      headers: { "Content-Type": "application/json" },
    });
  }
}
