/**
 * MiniMax provider — thin adapter over `@ai-sdk/openai-compatible`.
 *
 * Source of truth: SPEC §6 + the project config you gave me:
 *   base_url:  https://api.minimaxi.com/v1
 *   model:     MiniMax-M3  (context 512_000)
 *   auth:      Bearer token in Authorization header
 *
 * The Codex `wire_api: "responses"` flag in your config doesn't have a
 * direct counterpart in `@ai-sdk/openai-compatible` (which only speaks
 * Chat Completions today). The default Chat Completions endpoint at
 * `/v1/chat/completions` is what most OpenAI-compatible providers
 * expose and what this adapter targets. If MiniMax actually requires
 * the Responses API shape, we'll need to switch the underlying SDK or
 * add a `transformRequestBody` shim — flag it and we'll iterate.
 *
 * listModels() is the new bit: it actually hits GET /v1/models and
 * returns the IDs. The hardcoded `maxOutputTokens` / `capabilities`
 * come from a small local table keyed on model id, because the
 * OpenAI-compatible models endpoint is famously stingy with metadata
 * (just `id` + `object` + `owned_by`).
 */
import { createOpenAICompatible } from '@ai-sdk/openai-compatible'
import { streamText, stepCountIs, type LanguageModelUsage } from 'ai'
import type {
  ChatProvider,
  ModelInfo,
  StreamChatOpts,
  TokenUsage,
} from './interface'
import type { Cost, Message, StreamChunk } from '../types'

const PRICING_VERSION = '2026-06-08'

/**
 * Local metadata table for models we know about. `contextWindow`
 * comes from the project config; `maxOutputTokens` is a guess that
 * we'll refine as the provider publishes actual limits. Falls back
 * to safe defaults in fetchModelList() when the model id isn't
 * recognised.
 */
const KNOWN_MODEL_META: Record<
  string,
  { contextWindow: number; maxOutputTokens: number; capabilities?: string[] }
> = {
  'MiniMax-M3': {
    contextWindow: 512_000,
    maxOutputTokens: 16_000,
    capabilities: ['tools'],
  },
}

const DEFAULT_META = {
  contextWindow: 128_000,
  maxOutputTokens: 8_000,
}

// ---------- Message conversion ----------
//
// Chat Completions API accepts the same role/content shape we use
// internally, so messages pass through unchanged. Tool messages
// need a tiny rewrite (we store {toolCallId, toolName, content} in
// the internal Message; Chat Completions wants {role:'tool',
// tool_call_id, content}).

function toChatCompletionMessages(messages: Message[]): {
  role: 'user' | 'assistant' | 'tool' | 'system'
  content: string
  tool_call_id?: string
  tool_calls?: Array<{ id: string; type: 'function'; function: { name: string; arguments: string } }>
}[] {
  return messages.map((m) => {
    if (m.role === 'user') {
      return { role: 'user', content: m.content }
    }
    if (m.role === 'assistant') {
      const out: ReturnType<typeof toChatCompletionMessages>[number] = {
        role: 'assistant',
        content: m.content,
      }
      if (m.toolCalls && m.toolCalls.length > 0) {
        out.tool_calls = m.toolCalls.map((tc) => ({
          id: tc.id,
          type: 'function' as const,
          function: {
            name: tc.name,
            arguments: JSON.stringify(tc.args ?? {}),
          },
        }))
      }
      return out
    }
    // m.role === 'tool'
    return { role: 'tool', content: m.content, tool_call_id: m.toolCallId }
  })
}

// ---------- Pricing ----------

/**
 * Per-token USD pricing for MiniMax-M3 (placeholder; the provider
 * doesn't publish public rates — flag and we'll re-confirm with
 * the user's contract). We use 0 as the default; the chat UI
 * surfaces cost only when the user opts in.
 */
const PRICING: Record<string, { input: number; output: number }> = {
  'MiniMax-M3': { input: 0, output: 0 },
}

// ---------- Provider factory ----------

export type MiniMaxProviderOpts = {
  apiKey: string
  baseURL?: string
}

export function createMiniMaxProvider(opts: MiniMaxProviderOpts): ChatProvider {
  const provider = createOpenAICompatible({
    name: 'minimax',
    baseURL: opts.baseURL ?? 'https://api.minimaxi.com/v1',
    apiKey: opts.apiKey,
  })

  return {
    id: 'minimax',
    name: 'MiniMax',

    async *streamChat(s: StreamChatOpts): AsyncIterable<StreamChunk> {
      const result = streamText({
        model: provider(s.model),
        // The converted shape is structurally compatible with
        // ModelMessage but TS wants the literal role unions, so we
        // cast at the boundary rather than retype every branch.
        messages: toChatCompletionMessages(s.messages) as unknown as import('ai').ModelMessage[],
        system: s.system?.content,
        tools: s.tools,
        stopWhen: stepCountIs(s.maxSteps ?? 5),
        abortSignal: s.signal,
        ...(s.temperature !== undefined ? { temperature: s.temperature } : {}),
        ...(s.maxOutputTokens !== undefined
          ? { maxOutputTokens: s.maxOutputTokens }
          : {}),
      })

      let inputTokens = 0
      let outputTokens = 0
      let finishReason: 'stop' | 'tool-calls' | 'length' | 'error' = 'stop'

      for await (const part of result.fullStream) {
        switch (part.type) {
          case 'text-delta':
            yield { type: 'text-delta', delta: part.text }
            break
          case 'reasoning-delta':
            yield { type: 'reasoning-delta', delta: part.text }
            break
          case 'tool-input-start':
            yield {
              type: 'tool-call-start',
              toolCall: {
                id: part.id,
                name: part.toolName,
                args: {},
                status: 'pending',
              },
            }
            break
          case 'tool-input-delta':
            yield {
              type: 'tool-call-args',
              toolCallId: part.id,
              delta: part.delta,
            }
            break
          case 'tool-call':
            yield {
              type: 'tool-call-start',
              toolCall: {
                id: part.toolCallId,
                name: part.toolName,
                args: part.input,
                status: 'pending',
              },
            }
            break
          case 'tool-result':
            yield {
              type: 'tool-call-result',
              toolCallId: part.toolCallId,
              result: part.output,
            }
            break
          case 'tool-error':
            yield {
              type: 'tool-call-error',
              toolCallId: part.toolCallId,
              error:
                part.error instanceof Error
                  ? part.error.message
                  : String(part.error),
            }
            break
          case 'finish-step': {
            const u: LanguageModelUsage = part.usage
            inputTokens += u.inputTokens ?? 0
            outputTokens += u.outputTokens ?? 0
            break
          }
          case 'finish': {
            inputTokens = Math.max(inputTokens, part.totalUsage.inputTokens ?? 0)
            outputTokens = Math.max(outputTokens, part.totalUsage.outputTokens ?? 0)
            switch (part.finishReason) {
              case 'stop':
                finishReason = 'stop'
                break
              case 'tool-calls':
                finishReason = 'tool-calls'
                break
              case 'length':
                finishReason = 'length'
                break
              default:
                finishReason = 'stop'
            }
            break
          }
          case 'error':
            finishReason = 'error'
            yield {
              type: 'finish',
              reason: 'error',
              error:
                part.error instanceof Error ? part.error.message : String(part.error),
            }
            return
        }
      }

      yield { type: 'usage', inputTokens, outputTokens }
      yield { type: 'finish', reason: finishReason }
    },

    async listModels(listOpts: { apiKey: string; baseUrl?: string }): Promise<ModelInfo[]> {
      // The OpenAI-compatible /v1/models endpoint. Returns IDs only;
      // we merge in the local metadata table for context window etc.
      const base = (listOpts.baseUrl ?? 'https://api.minimaxi.com/v1').replace(
        /\/$/,
        '',
      )
      const url = `${base}/models`
      try {
        const r = await fetch(url, {
          headers: { Authorization: `Bearer ${listOpts.apiKey}` },
        })
        if (!r.ok) {
          throw new Error(`/v1/models returned ${r.status}`)
        }
        const body = (await r.json()) as {
          data?: Array<{ id: string; owned_by?: string }>
        }
        const ids = (body.data ?? []).map((m) => m.id)
        if (ids.length === 0) {
          throw new Error('/v1/models returned empty data array')
        }
        return ids.map((id) => modelInfoFromLocal(id))
      } catch (err) {
        // If the live fetch fails (network down, auth glitch, rate
        // limit), fall back to whatever we know locally so the
        // Settings panel isn't empty.
        console.warn('[minimax] listModels live fetch failed, using fallback:', err)
        return Object.keys(KNOWN_MODEL_META).map(modelInfoFromLocal)
      }
    },

    estimateCost(usage: TokenUsage, model: string): Cost {
      const p = PRICING[model]
      if (!p) {
        return { input: 0, output: 0, total: 0, source: 'tokens-only' }
      }
      const input = usage.inputTokens * p.input
      const output = usage.outputTokens * p.output
      return {
        input,
        output,
        total: input + output,
        source: 'estimated',
        pricingVersion: PRICING_VERSION,
      }
    },
  }
}

function modelInfoFromLocal(id: string): ModelInfo {
  const meta = KNOWN_MODEL_META[id] ?? DEFAULT_META
  return {
    id,
    displayName: id,
    contextWindow: meta.contextWindow,
    maxOutputTokens: meta.maxOutputTokens,
    capabilities: meta.capabilities,
  }
}
