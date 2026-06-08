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

export type ProviderId = 'anthropic' | 'openai' | 'openai-compatible'

export type ProviderSettings = {
  providerId: ProviderId
  apiKey: string
  /** Provider-scoped model id (e.g. 'claude-sonnet-4-5'). */
  model: string
  /** Only meaningful for openai-compatible. */
  baseUrl?: string
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
    case 'openai':
    case 'openai-compatible':
      // Phase 4 ships anthropic first; openai / openai-compatible are
      // Phase 7 work (SPEC §13 W25).
      throw new Error(
        `Provider "${settings.providerId}" is not yet implemented (Phase 7)`,
      )
    default: {
      const _exhaustive: never = settings.providerId
      throw new Error(`Unknown providerId: ${String(_exhaustive)}`)
    }
  }
}
