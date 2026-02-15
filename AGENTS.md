# Repository Guidelines

## Project Structure & Module Organization
- `src/index.ts` is the Cloudflare Worker entry point (HTTP, CORS, API key gate).
- `src/server.ts` wires MCP tool registration and dispatch.
- `src/tools/` contains domain handlers: `documents.ts` and `lists.ts`.
- `src/auth.ts`, `src/graph-client.ts`, and `src/middleware/auth.ts` handle Graph auth, API calls, and request validation.
- `src/types/` stores shared TypeScript types (`env.ts`, `models.ts`).
- Planning/design artifacts live in `PRD/` and `docs/`; CI deployment is in `.github/workflows/deploy.yml`.

## Build, Test, and Development Commands
- `npm install`: install dependencies.
- `npm run dev`: run local Worker with Wrangler.
- `npm run deploy`: deploy to Cloudflare Workers manually.
- `npm run tail`: stream Worker logs.
- `npm run cf-typegen`: regenerate Cloudflare type bindings.
- `npx tsc --noEmit`: strict type-check pass (recommended before PR).

## Coding Style & Naming Conventions
- Language: TypeScript (ES modules, `strict` mode enabled in `tsconfig.json`).
- Indentation: 2 spaces; keep trailing commas in multiline objects/arrays.
- Filenames use kebab-case (`graph-client.ts`); exported handlers use verb-first camelCase (`handleGetListItems`).
- Prefer small, focused functions with explicit return/error handling for MCP tool calls.
- Use existing comment style: short, intent-focused comments above non-obvious blocks.

## Testing Guidelines
- No automated test suite is configured yet; current validation is manual.
- Validate via `npm run dev` and an MCP client (for example Claude Desktop) against `/mcp`.
- For each change, test at least one success path and one failure path (auth failure, Graph API error, invalid input).
- Document manual test steps and results in the PR description.

## Commit & Pull Request Guidelines
- Follow existing commit style from history: imperative, concise subjects (for example `Fix list_files root path handling`).
- Keep commits scoped to one change area; avoid mixing refactors with feature work.
- PRs should include: purpose, affected MCP tools/endpoints, manual test evidence, and any required config/secret updates.
- Link related issue(s) when available and include logs/screenshots only when they clarify behavior.

## Security & Configuration Tips
- Never commit secrets; set `MCP_API_KEY`, Azure credentials, and `SHAREPOINT_SITE_URL` via Wrangler/Cloudflare secrets.
- Use the `GRAPH_TOKEN_CACHE` KV binding for token caching in deployed environments.
