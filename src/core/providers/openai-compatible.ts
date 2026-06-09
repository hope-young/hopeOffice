/**
 * Generic OpenAI-compatible provider — for self-hosted or
 * third-party endpoints that speak the Chat Completions API
 * (Ollama, LM Studio, vLLM, OpenRouter, …). Unlike `openai.ts`,
 * which uses the dedicated @ai-sdk/openai SDK, this delegates
 * to the shared `openai-compatible-chat` factory.
 *
 * The base URL is mandatory; the user picks it in Settings.
 */
import type { ModelInfo } from './interface'
import { createOpenAICompatibleChatProvider } from './openai-compatible-chat'

const KNOWN_MODELS: ModelInfo[] = [] // We trust the live fetch.

export type OpenAICompatibleProviderOpts = {
  apiKey: string
  /** Mandatory. The Settings panel's "Base URL" field feeds
   *  straight into here. */
  baseURL: string
}

export function createOpenAICompatibleProvider(
  opts: OpenAICompatibleProviderOpts,
) {
  if (!opts.baseURL) {
    throw new Error(
      'OpenAI-compatible provider requires a base URL. Set it in Settings.',
    )
  }
  return createOpenAICompatibleChatProvider({
    id: 'openai-compatible',
    name: 'OpenAI-compatible',
    baseURL: opts.baseURL,
    apiKey: opts.apiKey,
    providerName: 'openai-compatible',
    knownModels: KNOWN_MODELS,
    pricing: {}, // User self-hosts; cost is opaque.
    pricingVersion: '2026-06-08',
  })
}
