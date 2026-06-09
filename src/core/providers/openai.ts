/**
 * OpenAI provider — the official Chat Completions endpoint,
 * served by @ai-sdk/openai. Defaults to `https://api.openai.com/v1`;
 * the user can override baseURL in Settings to point at an Azure
 * OpenAI deployment or a local proxy.
 *
 * The stream / listModels / cost shape mirrors the
 * `openai-compatible-chat` factory; we duplicate the body here
 * (rather than reusing the factory) because @ai-sdk/openai and
 * @ai-sdk/openai-compatible are separate SDK entry points with
 * separate model registries — sharing code would force a runtime
 * import of `@ai-sdk/openai-compatible` from every install, even
 * users who only want the official OpenAI endpoint.
 */
import { createOpenAI } from '@ai-sdk/openai'
import { streamText, stepCountIs, type LanguageModelUsage } from 'ai'
import type {
  ChatProvider,
  ModelInfo,
  StreamChatOpts,
  TokenUsage,
} from './interface'
import type { Cost, Message, StreamChunk } from '../types'

const PRICING_VERSION = '2026-06-08'

/** Curated model metadata. /v1/models is fetched live too, but
 *  without a per-model context window — we merge the two sources
 *  with the local table taking precedence. */
const KNOWN_MODELS: ModelInfo[] = [
  {
    id: 'gpt-4o',
    displayName: 'GPT-4o',
    contextWindow: 128_000,
    maxOutputTokens: 16_000,
    capabilities: ['vision', 'tools'],
  },
  {
    id: 'gpt-4o-mini',
    displayName: 'GPT-4o mini',
    contextWindow: 128_000,
    maxOutputTokens: 16_000,
    capabilities: ['vision', 'tools'],
  },
  {
    id: 'o3-mini',
    displayName: 'o3-mini',
    contextWindow: 200_000,
    maxOutputTokens: 100_000,
    capabilities: ['reasoning', 'tools'],
  },
]

/** Public OpenAI pricing, per 1M tokens (USD). 1e-6 = $1/MTok. */
const PRICING: Record<string, { input: number; output: number }> = {
  'gpt-4o': { input: 2.5e-6, output: 10e-6 },
  'gpt-4o-mini': { input: 0.15e-6, output: 0.6e-6 },
  'o3-mini': { input: 1.1e-6, output: 4.4e-6 },
}

export type OpenAIProviderOpts = {
  apiKey: string
  /** Optional — defaults to api.openai.com. Point this at an
   *  Azure OpenAI deployment or a proxy. */
  baseURL?: string
}

export function createOpenAIProvider(opts: OpenAIProviderOpts): ChatProvider {
  const provider = createOpenAI({
    apiKey: opts.apiKey,
    ...(opts.baseURL ? { baseURL: opts.baseURL } : {}),
  })

  return {
    id: 'openai',
    name: 'OpenAI',

    async *streamChat(s: StreamChatOpts): AsyncIterable<StreamChunk> {
      const result = streamText({
        model: provider(s.model),
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

    async listModels(listOpts: {
      apiKey: string
      baseUrl?: string
    }): Promise<ModelInfo[]> {
      const base = (listOpts.baseUrl ?? 'https://api.openai.com/v1').replace(/\/$/, '')
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
        console.warn('[openai] listModels live fetch failed, using fallback:', err)
        return KNOWN_MODELS
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

function modelInfoFromLocal(id: string): ModelInfo {
  const found = KNOWN_MODELS.find((m) => m.id === id)
  if (found) return found
  return {
    id,
    displayName: id,
    contextWindow: 128_000,
    maxOutputTokens: 8_000,
  }
}
