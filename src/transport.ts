// Custom MCP Transport for Cloudflare Workers (Web API compatible)

import { Transport } from "@modelcontextprotocol/sdk/shared/transport.js";
import { JSONRPCMessage } from "@modelcontextprotocol/sdk/types.js";

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

    const response = await new Promise<JSONRPCMessage>((resolve) => {
      this.responseResolve = resolve;
      this.onmessage?.(body);
    });

    return new Response(JSON.stringify(response), {
      headers: { "Content-Type": "application/json" },
    });
  }
}
