/**
 * Provider registry — the single place UI code goes to get a
 * ChatProvider. Source of truth: SPEC §6 (registry).
 *
 * The Settings panel owns the `ProviderSettings` object; UI hands
 * that object to `createProvider(settings)` and gets back a ready
 * ChatProvider. UI never imports a specific provider module.
 */
import type { ChatProvider } from './interface'
import { createAnthropicProvider } from './anthropic'
import { createMiniMaxProvider } from './minimax'

export type ProviderId = 'anthropic' | 'minimax' | 'openai' | 'openai-compatible'

export type ProviderSettings = {
  providerId: ProviderId
  apiKey: string
  /** Provider-scoped model id (e.g. 'claude-sonnet-4-5'). */
  model: string
  /** Only meaningful for openai-compatible / minimax with a custom endpoint. */
  baseUrl?: string
}

export type ListModelsOpts = {
  providerId: ProviderId
  apiKey: string
  baseUrl?: string
}

/**
 * Thin pass-through to ChatProvider.listModels so the Settings UI
 * can populate the model dropdown without knowing which concrete
 * provider to import. Throws for providerIds that don't have a
 * listModels implementation yet (see createProvider).
 */
export async function listModelsForProvider(
  opts: ListModelsOpts,
): Promise<import('./interface').ModelInfo[]> {
  const provider = createProvider({
    providerId: opts.providerId,
    apiKey: opts.apiKey,
    model: '',
    ...(opts.baseUrl ? { baseUrl: opts.baseUrl } : {}),
  })
  return provider.listModels({ apiKey: opts.apiKey, ...(opts.baseUrl ? { baseUrl: opts.baseUrl } : {}) })
}

/**
 * Build a ChatProvider from the user's settings. Throws when the
 * `providerId` is unknown (UI side should guard against that, but we
 * keep this strict rather than silently falling back to a stub).
 */
export function createProvider(settings: ProviderSettings): ChatProvider {
  switch (settings.providerId) {
    case 'anthropic':
      return createAnthropicProvider({ apiKey: settings.apiKey })
    case 'minimax':
      return createMiniMaxProvider({
        apiKey: settings.apiKey,
        ...(settings.baseUrl ? { baseURL: settings.baseUrl } : {}),
      })
    case 'openai':
    case 'openai-compatible':
      // Phase 7 work (SPEC §13 W25). Generic OpenAI / OpenAI-compatible
      // adapters can share most of minimax.ts's code — refactor that
      // into a shared base when we add them.
      throw new Error(
        `Provider "${settings.providerId}" is not yet implemented (Phase 7)`,
      )
    default: {
      const _exhaustive: never = settings.providerId
      throw new Error(`Unknown providerId: ${String(_exhaustive)}`)
    }
  }
}
