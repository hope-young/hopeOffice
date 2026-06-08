/**
 * ChatProvider — the abstraction every LLM backend implements.
 *
 * Source of truth: SPEC §6. The UI and the orchestrator talk to this
 * interface only; they never see provider SDK names.
 *
 * No DOM, no Office, no React imports allowed in this file.
 */
import type { Cost, Message, StreamChunk } from '../types'

// ---------- Tool / schema types ----------

/**
 * A JSON Schema fragment we pass to the LLM. We only type the fields we
 * actually use; providers can pass through SDK-specific extras as
 * `unknown` (Anthropic / OpenAI both tolerate a superset).
 */
export type JsonSchema = {
  type: 'object'
  properties: Record<string, unknown>
  required?: string[]
  additionalProperties?: boolean
  [key: `x-${string}`]: unknown
}

/**
 * A tool definition as the LLM sees it. The provider adapter turns
 * this into the SDK-specific shape (Anthropic `tools`, OpenAI
 * `tools`, OpenAI-compatible `tools`, …).
 *
 * By convention the orchestrator registers a single `execute_skill`
 * tool whose `parameters` is a union of every Skill's zod schema
 * (see SPEC_DETAILS §6 "Calling convention"). MCP tools are added
 * alongside in Phase 10.
 */
export type ToolDef = {
  name: string
  description: string
  parameters: JsonSchema
}

// ---------- Usage / pricing ----------

export type TokenUsage = {
  inputTokens: number
  outputTokens: number
  /** Cached input tokens, when the provider breaks them out
   *  (Anthropic does; OpenAI / OpenAI-compatible do not). */
  cachedInputTokens?: number
  /** Provider-specific raw usage blob, for debugging / cost
   *  reconciliation when the typed fields are insufficient. */
  raw?: unknown
}

// ---------- Model listing ----------

export type ModelInfo = {
  /** Provider-scoped identifier (`claude-opus-4-7`, `gpt-5`, …). */
  id: string
  displayName: string
  /** Max input tokens (prompt side). */
  contextWindow: number
  /** Max output tokens (completion side). */
  maxOutputTokens: number
  /** Free-form capability flags the UI can show as a chip:
   *  `"vision"`, `"tools"`, `"reasoning"`, … */
  capabilities?: string[]
}

// ---------- streamChat options ----------

export type StreamChatOpts = {
  messages: Message[]
  tools: ToolDef[]
  /** Provider-scoped model id. */
  model: string
  /** Cancelled when the user hits Stop or the conversation is closed. */
  signal: AbortSignal
  /** Optional sampling hints. Providers ignore fields they don't support
   *  rather than throwing — the user-facing Settings panel only
   *  exposes params that all three providers accept. */
  temperature?: number
  maxOutputTokens?: number
}

// ---------- The interface ----------

export interface ChatProvider {
  /** Stable identifier used in settings + URLs:
   *  `anthropic` | `openai` | `openai-compatible`. */
  id: string
  /** Human-readable provider name shown in the Settings dropdown. */
  name: string
  streamChat(opts: StreamChatOpts): AsyncIterable<StreamChunk>
  listModels(opts: { apiKey: string; baseUrl?: string }): Promise<ModelInfo[]>
  estimateCost(usage: TokenUsage, model: string): Cost
}
