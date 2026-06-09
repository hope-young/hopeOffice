/**
 * Shared types used by both `core/` (pure logic) and `taskpane/` (React UI).
 *
 * Source of truth: SPEC_DETAILS.md §1.
 * No DOM, no Office, no React imports allowed in this file.
 */

// ---------- Host ----------

export type HostKind = 'word' | 'excel' | 'powerpoint' | 'unsupported'

// ---------- Conversation ----------

export type Role = 'user' | 'assistant' | 'tool'

export type Message =
  | { role: 'user'; content: string }
  | {
      role: 'assistant'
      content: string
      /** Reasoning model output (e.g. MiniMax M3's `<think>` block).
       *  Stored separately so the UI can collapse it. */
      reasoning?: string
      toolCalls?: ToolCall[]
    }
  | { role: 'tool'; toolCallId: string; toolName: string; content: string }

// ---------- Streaming ----------

/** One chunk emitted by a ChatProvider's `streamChat`. */
export type StreamChunk =
  | { type: 'text-delta'; delta: string }
  | { type: 'reasoning-delta'; delta: string } // thinking models
  | { type: 'tool-call-start'; toolCall: ToolCall }
  | { type: 'tool-call-args'; toolCallId: string; delta: unknown } // JSON delta
  | { type: 'tool-call-result'; toolCallId: string; result: unknown }
  | { type: 'tool-call-error'; toolCallId: string; error: string }
  | {
      type: 'finish'
      reason: 'stop' | 'tool-calls' | 'length' | 'error'
      error?: string
    }
  | { type: 'usage'; inputTokens: number; outputTokens: number }

// ---------- Tool call ----------

export type ToolCall = {
  /** Unique within a single conversation. */
  id: string
  /** 'execute_skill' | 'execute_code' | '<mcp-tool-name>' */
  name: string
  /** Validated by Zod before execution. */
  args: unknown
  result?: unknown
  status: 'pending' | 'running' | 'success' | 'error'
  error?: { message: string; stack?: string }
  startedAt?: number
  completedAt?: number
}

// ---------- Cost ----------

export type CostSource =
  | 'gateway-exact' // Vercel AI Gateway
  | 'openrouter-exact'
  | 'estimated' // local price table
  | 'tokens-only' // no price table
  | 'local-free' // Ollama / LM Studio

export type Cost = {
  input: number
  output: number
  total: number
  source: CostSource
  /** Required when `source === 'estimated'`. */
  pricingVersion?: string
}

// ---------- Agent state (SPEC §5) ----------

export type AgentState = {
  messages: Message[]
  /** Current streaming text delta. */
  draft: string
  /** Streaming reasoning delta (collapsed in the UI; committed to
   *  `messages[last].reasoning` on stream-end). */
  draftReasoning: string
  toolCalls: ToolCall[]
  status: 'idle' | 'streaming' | 'awaiting-approval' | 'executing' | 'error'
  cost: Cost
  approval?: { code: string; resolve: (ok: boolean) => void }
  error?: Error
}

export type AgentEvent =
  | { type: 'user-send'; text: string }
  | { type: 'stream-token'; token: string }
  /** A reasoning-model delta (separate from `stream-token`) so the
   *  UI can render `<think>` blocks collapsed. */
  | { type: 'reasoning-delta'; delta: string }
  | { type: 'stream-end' }
  | { type: 'tool-call-start'; toolCall: ToolCall }
  | { type: 'tool-call-result'; toolCallId: string; result: unknown }
  | { type: 'tool-call-error'; toolCallId: string; error: string }
  | { type: 'approval-requested'; code: string }
  | { type: 'approval-decided'; approved: boolean }
  | { type: 'error'; error: Error }
  | { type: 'reset' }
  /** Internal: replace the messages array wholesale. Used by the
   *  history-restore path on app start. */
  | { type: 'restore-messages'; messages: Message[] }
