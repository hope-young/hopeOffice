/**
 * ChatProvider — the abstraction every LLM backend implements.
 *
 * Source of truth: SPEC §6. The UI and the orchestrator talk to this
 * interface only; they never see provider SDK names.
 *
 * No DOM, no Office, no React imports allowed in this file.
 */
import type { Tool } from 'ai'
import type { Cost, Message, StreamChunk } from '../types'

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

/**
 * A system message to prepend to the conversation. The provider
 * adapter turns this into the SDK-specific shape (Anthropic
 * `system`, OpenAI `messages[0].role:'system'`, …).
 */
export type SystemMessage = {
  content: string
}

export type StreamChatOpts = {
  messages: Message[]
  /**
   * AI SDK 5 tool definitions. The orchestrator builds this from
   * the host's skill registry. Providers pass it through to
   * `streamText({ tools })` unchanged.
   */
  tools: Record<string, Tool>
  /** Optional system prompt — orchestrator builds this from the
   *  skill registry + host. */
  system?: SystemMessage
  /** Provider-scoped model id. */
  model: string
  /** Cancelled when the user hits Stop or the conversation is closed. */
  signal: AbortSignal
  /** Optional sampling hints. Providers ignore fields they don't support
   *  rather than throwing — the user-facing Settings panel only
   *  exposes params that all three providers accept. */
  temperature?: number
  maxOutputTokens?: number
  /**
   * Max number of LLM round-trips in a single user turn (each
   * round-trip can call tools). Default 1 (no multi-step). Phase 5
   * bumps to ~5 so the LLM can chain skills in one turn.
   */
  maxSteps?: number
}

// ---------- The interface ----------

export interface ChatProvider {
  /** Stable identifier used in settings + URLs:
   *  `anthropic` | `minimax` | `openai` | `openai-compatible`. */
  id: string
  /** Human-readable provider name shown in the Settings dropdown. */
  name: string
  streamChat(opts: StreamChatOpts): AsyncIterable<StreamChunk>
  listModels(opts: { apiKey: string; baseUrl?: string }): Promise<ModelInfo[]>
  estimateCost(usage: TokenUsage, model: string): Cost
}
