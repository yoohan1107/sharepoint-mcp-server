// API Key validation middleware

import { Env } from "../types/env";

/**
 * Validate MCP API Key from Authorization header
 * Uses constant-time comparison to prevent timing attacks
 */
export async function validateAPIKey(
  request: Request,
  env: Env
): Promise<boolean> {
  const authHeader = request.headers.get("Authorization");

  if (!authHeader) {
    return false;
  }

  // Expected format: "Bearer {API_KEY}"
  const [scheme, token] = authHeader.split(" ");

  if (scheme !== "Bearer" || !token) {
    return false;
  }

  // Constant-time comparison
  return safeCompare(token, env.MCP_API_KEY);
}

/**
 * Constant-time string comparison to prevent timing attacks
 */
function safeCompare(a: string, b: string): boolean {
  if (a.length !== b.length) {
    return false;
  }

  let result = 0;
  for (let i = 0; i < a.length; i++) {
    result |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }

  return result === 0;
}
