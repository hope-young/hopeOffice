/**
 * MCP (Model Context Protocol) client wrapper.
 *
 * Source of truth: SPEC §13 W27 + SPEC_DETAILS §3. Per the spec,
 * the registry MAY register MCP-server tools at startup; their
 * behaviour is identical to skills from the LLM's perspective
 * (Zod-validated input, async execute, streaming results). We
 * expose MCP server tools to the LLM the same way we expose
 * built-in skills — through the AI SDK 5 `tools` record on
 * `streamText`.
 *
 * Per SPEC §13: "MCP server 工具由 orchestrator 启动时注册".
 * We take "startup" loosely — clients are lazy: a per-server
 * MCP client is created on first `getMcpTools()` call for that
 * server and cached for the lifetime of the page. The cache
 * survives across `send` rounds so we don't pay the
 * HTTP/SSE handshake cost on every turn.
 *
 * Source of truth: SPEC §10. SPEC §2 says "不接 stdio MCP
 * (只 HTTP/SSE)" — we honour that and only accept the two
 * streamable transports.
 */
import { createMCPClient, type MCPClient } from '@ai-sdk/mcp'
import type { Tool } from 'ai'

/** Configuration for a single MCP server, as stored in Settings. */
export type McpServerConfig = {
  /** Stable id (also used as the tool-name prefix so the LLM can
   *  tell servers apart: `<serverName>__<toolName>`). */
  name: string
  /** Full URL to the MCP endpoint, e.g.
   *  `https://my-mcp.example.com/mcp` (HTTP) or `…/sse` (SSE). */
  url: string
  /** Transport flavour. */
  transport: 'http' | 'sse'
}

// ---------- Client cache ----------

type CacheEntry = {
  client: MCPClient
  tools: Record<string, Tool>
}

const cache = new Map<string, CacheEntry>()

function cacheKey(s: McpServerConfig): string {
  return `${s.name}|${s.url}|${s.transport}`
}

function buildTransport(s: McpServerConfig): { transport: { type: 'http' | 'sse'; url: string } } {
  // @ai-sdk/mcp 1.0.46's `createMCPClient` takes a
  // `{ transport: MCPTransportConfig }` shape — not a flat
  // `{ type, url }`. The `'transport'` name in our user-facing
  // `McpServerConfig` maps 1:1 to the MCP spec's transport
  // 'type' value.
  return { transport: { type: s.transport, url: s.url } }
}

/** Get (and lazily create) a cached MCP client + its tool set. */
async function getOrCreate(s: McpServerConfig): Promise<CacheEntry> {
  const key = cacheKey(s)
  const hit = cache.get(key)
  if (hit) return hit
  const client = await createMCPClient(buildTransport(s))
  // Cast: @ai-sdk/mcp 1.0.46's `tools()` returns an MCP-flavored
  // ToolSet (V3 protocol). We use it as a generic `Record<string,
  // Tool>` at the boundary; the consumer (streamText) accepts
  // either shape.
  const tools: Record<string, Tool> = (await client.tools()) as unknown as Record<string, Tool>
  const entry: CacheEntry = { client, tools }
  cache.set(key, entry)
  return entry
}

/**
 * Drop a server from the cache (used by the Settings UI when
 * the user edits or removes a server — forces the next `send`
 * to reconnect with the new config).
 */
export function invalidateMcpServer(name: string): void {
  for (const k of Array.from(cache.keys())) {
    if (k.startsWith(`${name}|`)) cache.delete(k)
  }
}

/**
 * Drop every cached MCP client. Called on hard reset (test
 * teardown, full page reload after Settings wipe).
 */
export function clearMcpCache(): void {
  for (const entry of cache.values()) {
    try {
      void entry.client.close()
    } catch {
      // Best-effort close; ignore failures.
    }
  }
  cache.clear()
}

// ---------- Public API ----------

/**
 * Fetch and merge tools across every reachable MCP server in
 * `servers`. Each tool is keyed `<serverName>__<toolName>` so
 * two servers can expose a same-named tool without colliding.
 *
 * Errors per server are caught and logged: a single misbehaving
 * server doesn't take down the rest of the tool surface. The
 * orchestrator treats an empty `tools` map as "no MCP" (same
 * as zero configured).
 */
export async function getMcpTools(
  servers: McpServerConfig[],
): Promise<Record<string, Tool>> {
  const out: Record<string, Tool> = {}
  if (servers.length === 0) return out
  const results = await Promise.allSettled(
    servers.map(async (s) => {
      const entry = await getOrCreate(s)
      return { name: s.name, tools: entry.tools }
    }),
  )
  for (let i = 0; i < results.length; i++) {
    const r = results[i]!
    if (r.status === 'fulfilled') {
      const { name, tools } = r.value
      for (const [toolName, tool] of Object.entries(tools)) {
        out[`${name}__${toolName}`] = tool
      }
    } else {
      const s = servers[i]!
      console.warn(
        `[mcp] failed to fetch tools from server "${s.name}":`,
        r.reason,
      )
    }
  }
  return out
}

// ---------- Test helpers ----------

/** Reset module state. Used by tests. */
export function _resetForTests(): void {
  clearMcpCache()
}
