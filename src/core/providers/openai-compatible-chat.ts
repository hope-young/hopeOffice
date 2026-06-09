/**
 * Shared OpenAI-compatible chat-completions adapter.
 *
 * The MiniMax provider, the generic OpenAI-compatible provider
 * (Ollama, LM Studio, vLLM, OpenRouter, …), and any future
 * Chat-Completions-shaped endpoint share this exact code path
 * through `@ai-sdk/openai-compatible`. We factor the body into
 * a factory so each provider entry point is just config + a known
 * model table.
 *
 * What this does NOT do:
 *   - Responses-API mode. `@ai-sdk/openai-compatible` only speaks
 *     `/v1/chat/completions` today. The OpenAI official provider
 *     can opt into Responses via `createOpenAI({ useResponses: true })`,
 *     which is a separate file (see `openai.ts`).
 *   - Server-side function-calling extension. We translate to /
 *     from the AI SDK 5 TextStreamPart shape; provider-specific
 *     tool-calling extensions (e.g. Anthropic's prompt caching)
 *     would need a custom adapter.
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

export type OpenAICompatibleChatProviderOpts = {
  /** Stable id used by the registry (e.g. 'minimax',
   *  'openai-compatible'). */
  id: string
  /** Human-readable name shown in the Settings dropdown. */
  name: string
  /** The base URL (without trailing slash). */
  baseURL: string
  /** Bearer token. */
  apiKey: string
  /** Adapter name passed to createOpenAICompatible ({ name }). */
  providerName: string
  /** Curated model metadata table. `id` is the provider-scoped
   *  model id (e.g. 'MiniMax-M3', 'gpt-4o'). Used as a fallback
   *  when the /v1/models fetch fails or returns empty. */
  knownModels: ModelInfo[]
  /** Optional per-token USD pricing (input/output). Empty object
   *  means cost is reported as `tokens-only` until the user
   *  fills in a real table. */
  pricing: Record<string, { input: number; output: number }>
  /** Pricing version tag for the `estimated` cost source. */
  pricingVersion: string
}

export function createOpenAICompatibleChatProvider(
  opts: OpenAICompatibleChatProviderOpts,
): ChatProvider {
  const provider = createOpenAICompatible({
    name: opts.providerName,
    baseURL: opts.baseURL,
    apiKey: opts.apiKey,
  })

  return {
    id: opts.id,
    name: opts.name,

    async *streamChat(s: StreamChatOpts): AsyncIterable<StreamChunk> {
      const result = streamText({
        model: provider(s.model),
        // The converted shape is structurally compatible with
        // ModelMessage but TS wants the literal role unions, so
        // we cast at the boundary rather than retype every branch.
        messages: toChatCompletionMessages(s.messages) as unknown as import('ai').ModelMessage[],
        system: s.system?.content,
        tools: s.tools as any,
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

    async listModels(listOpts: {
      apiKey: string
      baseUrl?: string
    }): Promise<ModelInfo[]> {
      // The OpenAI-compatible /v1/models endpoint. Returns IDs only;
      // we merge in the local metadata table for context window etc.
      // — the OpenAI models endpoint is famously stingy with
      // metadata (just `id` + `object` + `owned_by`).
      const base = (listOpts.baseUrl ?? opts.baseURL).replace(/\/$/, '')
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
        return ids.map((id) => modelInfoFromLocal(id, opts.knownModels))
      } catch (err) {
        // If the live fetch fails (network down, auth glitch, rate
        // limit), fall back to whatever we know locally so the
        // Settings panel isn't empty during a transient blip.
        console.warn(`[${opts.id}] listModels live fetch failed, using fallback:`, err)
        return opts.knownModels
      }
    },

    estimateCost(usage: TokenUsage, model: string): Cost {
      const p = opts.pricing[model]
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
        pricingVersion: opts.pricingVersion,
      }
    },
  }
}

// ---------- helpers ----------

function toChatCompletionMessages(
  messages: Message[],
): {
  role: 'user' | 'assistant' | 'tool' | 'system'
  content: string
  tool_call_id?: string
  tool_calls?: Array<{
    id: string
    type: 'function'
    function: { name: string; arguments: string }
  }>
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

function modelInfoFromLocal(id: string, known: ModelInfo[]): ModelInfo {
  const found = known.find((m) => m.id === id)
  if (found) return found
  return {
    id,
    displayName: id,
    contextWindow: 128_000,
    maxOutputTokens: 8_000,
  }
}
