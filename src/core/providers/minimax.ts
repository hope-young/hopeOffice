/**
 * MiniMax provider — thin configuration wrapper over the shared
 * OpenAI-compatible chat adapter. Source of truth: the project
 * config the user gave me:
 *   base_url:  https://api.minimaxi.com/v1
 *   model:     MiniMax-M3  (context 512_000)
 *   auth:      Bearer token in Authorization header
 *
 * The actual stream / listModels / estimateCost logic lives in
 * `openai-compatible-chat.ts` so we don't duplicate 200+ lines
 * per provider. This file is configuration only.
 */
import type { ModelInfo } from './interface'
import { createOpenAICompatibleChatProvider } from './openai-compatible-chat'

/**
 * Curated list of MiniMax models. The /v1/models endpoint
 * returns just IDs; we merge in the context window + capabilities
 * locally. Per-token USD pricing is left at 0 — MiniMax doesn't
 * publish public rates — so the cost UI shows 0 until the user
 * fills in real numbers.
 */
const KNOWN_MODELS: ModelInfo[] = [
  {
    id: 'MiniMax-M3',
    displayName: 'MiniMax-M3',
    contextWindow: 512_000,
    maxOutputTokens: 16_000,
    capabilities: ['tools'],
  },
]

const PRICING: Record<string, { input: number; output: number }> = {
  'MiniMax-M3': { input: 0, output: 0 },
}

export type MiniMaxProviderOpts = {
  apiKey: string
  baseURL?: string
}

export function createMiniMaxProvider(opts: MiniMaxProviderOpts) {
  return createOpenAICompatibleChatProvider({
    id: 'minimax',
    name: 'MiniMax',
    baseURL: opts.baseURL ?? 'https://api.minimaxi.com/v1',
    apiKey: opts.apiKey,
    providerName: 'minimax',
    knownModels: KNOWN_MODELS,
    pricing: PRICING,
    pricingVersion: '2026-06-08',
  })
}
