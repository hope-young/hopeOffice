/**
 * Anthropic provider — thin adapter around `@ai-sdk/anthropic` +
 * Vercel AI SDK 5's `streamText`. Source of truth: SPEC §6 +
 * SPEC_DETAILS §6.
 *
 * The adapter is intentionally thin: it translates between our
 * ChatProvider surface (StreamChunk / ToolDef / Message) and the AI
 * SDK surface (LanguageModel / CoreMessage / TextStreamPart). UI and
 * orchestrator code never imports `@ai-sdk/anthropic` directly.
 *
 * What this Phase-4 version does NOT do:
 *   - Forward `tools` to the model. Tool calls land in Phase 5
 *     (skill registry) when we have real Skills to advertise.
 *   - Hit `/v1/models`. Anthropic's models endpoint returns a sparse
 *     list; the curated KNOWN_MODELS below covers what the Settings
 *     dropdown needs.
 */
import { createAnthropic } from '@ai-sdk/anthropic'
import { streamText, stepCountIs, type ModelMessage } from 'ai'
import type {
  ChatProvider,
  ModelInfo,
  StreamChatOpts,
  TokenUsage,
} from './interface'
import type { Cost, Message, StreamChunk } from '../types'

const PRICING_VERSION = '2026-06-08'

/**
 * Curated list of Anthropic models. Updated by hand when a new tier
 * ships. Prices are per-token USD; cached input tokens are billed at
 * ~10% of the standard input rate (Anthropic's published caching
 * discount), reflected in the math below.
 */
const KNOWN_MODELS: ModelInfo[] = [
  {
    id: 'claude-opus-4-7',
    displayName: 'Claude Opus 4.7',
    contextWindow: 200_000,
    maxOutputTokens: 32_000,
    capabilities: ['vision', 'tools', 'reasoning'],
  },
  {
    id: 'claude-sonnet-4-5',
    displayName: 'Claude Sonnet 4.5',
    contextWindow: 200_000,
    maxOutputTokens: 16_000,
    capabilities: ['vision', 'tools'],
  },
  {
    id: 'claude-haiku-4-5',
    displayName: 'Claude Haiku 4.5',
    contextWindow: 200_000,
    maxOutputTokens: 8_000,
    capabilities: ['vision', 'tools'],
  },
]

const PRICING: Record<string, { input: number; output: number }> = {
  'claude-opus-4-7': { input: 15e-6, output: 75e-6 },
  'claude-sonnet-4-5': { input: 3e-6, output: 15e-6 },
  'claude-haiku-4-5': { input: 1e-6, output: 5e-6 },
}

// ---------- Message conversion ----------

function toCoreMessages(messages: Message[]): ModelMessage[] {
  const out: ModelMessage[] = []
  for (const m of messages) {
    if (m.role === 'user') {
      out.push({ role: 'user', content: m.content })
      continue
    }
    if (m.role === 'assistant') {
      const am: ModelMessage = { role: 'assistant', content: m.content }
      if (m.toolCalls && m.toolCalls.length > 0) {
        // The AI SDK 5 CoreAssistantMessage shape; untyped here because
        // the assistant-toolCalls field name varies between SDK 4/5.
        ;(am as unknown as { toolCalls: unknown[] }).toolCalls = m.toolCalls.map(
          (tc) => ({
            toolCallId: tc.id,
            toolName: tc.name,
            input: tc.args,
          }),
        )
      }
      out.push(am)
      continue
    }
    // m.role === 'tool'
    out.push({
      role: 'tool',
      content: [
        {
          type: 'tool-result',
          toolCallId: m.toolCallId,
          toolName: m.toolName,
          output: { type: 'text', value: m.content },
        },
      ],
    })
  }
  return out
}

// ---------- Provider factory ----------

export type AnthropicProviderOpts = {
  apiKey: string
}

export function createAnthropicProvider(opts: AnthropicProviderOpts): ChatProvider {
  const anthropic = createAnthropic({ apiKey: opts.apiKey })

  return {
    id: 'anthropic',
    name: 'Anthropic',

    async *streamChat(s: StreamChatOpts): AsyncIterable<StreamChunk> {
      // Phase 5: forward tools + system to the AI SDK. The SDK runs
      // skill.execute locally, handles the multi-step loop via
      // `stopWhen: stepCountIs(maxSteps)`, and emits a typed
      // `fullStream` we re-shape into our `StreamChunk` vocabulary.
      const result = streamText({
        model: anthropic(s.model),
        system: s.system?.content,
        messages: toCoreMessages(s.messages),
        tools: s.tools,
        stopWhen: stepCountIs(s.maxSteps ?? 5),
        abortSignal: s.signal,
        ...(s.temperature !== undefined ? { temperature: s.temperature } : {}),
        ...(s.maxOutputTokens !== undefined
          ? { maxOutputTokens: s.maxOutputTokens }
          : {}),
      })

      // AI SDK 5 emits a typed fullStream; we walk it and re-emit our
      // own StreamChunk shape. text-delta / reasoning-delta / tool-*
      // / finish / usage all line up. The 'error' chunk becomes a
      // finish with reason='error' and the error message.
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
            // StreamChunk.tool-call-start takes a complete ToolCall.
            // We build it with an empty `args` placeholder; the
            // subsequent tool-input-delta events fill it in.
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
            // The AI SDK fires this with the *complete* validated
            // input. We dispatch a fresh tool-call-start so the
            // orchestrator replaces the placeholder.
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
          case 'finish-step':
            // Per-step usage; we sum across steps for the final usage
            // chunk.
            inputTokens += part.usage.inputTokens ?? 0
            outputTokens += part.usage.outputTokens ?? 0
            break
          case 'finish':
            inputTokens += part.totalUsage.inputTokens ?? 0
            outputTokens += part.totalUsage.outputTokens ?? 0
            // part.totalUsage sometimes only reports step totals at
            // finish-step time; the `finish` part carries the final
            // accumulator in some SDK versions. Use whichever is
            // larger to be safe.
            if (part.totalUsage.inputTokens != null) {
              inputTokens = Math.max(inputTokens, part.totalUsage.inputTokens)
            }
            if (part.totalUsage.outputTokens != null) {
              outputTokens = Math.max(outputTokens, part.totalUsage.outputTokens)
            }
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
          case 'error':
            finishReason = 'error'
            yield {
              type: 'finish',
              reason: 'error',
              error: part.error instanceof Error ? part.error.message : String(part.error),
            }
            return
        }
      }

      yield { type: 'usage', inputTokens, outputTokens }
      yield { type: 'finish', reason: finishReason }
    },

    async listModels(_listOpts: { apiKey: string; baseUrl?: string }): Promise<ModelInfo[]> {
      // See note above; we don't hit /v1/models.
      return KNOWN_MODELS
    },

    estimateCost(usage: TokenUsage, model: string): Cost {
      const p = PRICING[model]
      if (!p) {
        return { input: 0, output: 0, total: 0, source: 'tokens-only' }
      }
      const cached = usage.cachedInputTokens ?? 0
      const billableInput = Math.max(0, usage.inputTokens - cached)
      // Cached input billed at ~10% of standard rate (Anthropic's
      // published caching discount).
      const input = billableInput * p.input + cached * p.input * 0.1
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
