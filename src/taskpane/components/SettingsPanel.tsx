/**
 * SettingsPanel — provider, API key, model, base URL. Persisted
 * via the Zustand settings store (localStorage).
 *
 * Model list is now dynamic: it calls
 * `listModelsForProvider({providerId, apiKey, baseUrl})` whenever
 * the relevant inputs change, and falls back to a small local list
 * when the network call fails. Anthropic and OpenAI (when added in
 * Phase 7) keep their hardcoded lists — /v1/models is unreliable
 * on those.
 */
import { useEffect, useState } from 'react'
import type { ProviderId } from '@core/providers/registry'
import { listModelsForProvider } from '@core/providers/registry'
import type { ModelInfo } from '@core/providers/interface'
import { useSettingsStore } from '../store/settings'

/** Curated list shown when the live fetch is unavailable (Anthropic). */
const ANTHROPIC_MODELS: ModelInfo[] = [
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

/** Default model id per provider — used on first load. */
const DEFAULT_MODEL: Record<ProviderId, string> = {
  anthropic: 'claude-sonnet-4-5',
  minimax: 'MiniMax-M3',
  openai: 'gpt-4o',
  'openai-compatible': '',
}

/** Whether the provider has a customisable base URL. */
const HAS_BASE_URL: Record<ProviderId, boolean> = {
  anthropic: false,
  minimax: true,
  openai: true,
  'openai-compatible': true,
}

/** Whether the provider's model list is fetched live. */
const FETCHES_LIVE_MODELS: Record<ProviderId, boolean> = {
  anthropic: false,
  minimax: true,
  openai: true,
  'openai-compatible': true,
}

export function SettingsPanel() {
  const providerId = useSettingsStore((s) => s.providerId)
  const apiKey = useSettingsStore((s) => s.apiKey)
  const model = useSettingsStore((s) => s.model)
  const baseUrl = useSettingsStore((s) => s.baseUrl)
  const mcpServers = useSettingsStore((s) => s.mcpServers)

  const setProviderId = useSettingsStore((s) => s.setProviderId)
  const setApiKey = useSettingsStore((s) => s.setApiKey)
  const setModel = useSettingsStore((s) => s.setModel)
  const setBaseUrl = useSettingsStore((s) => s.setBaseUrl)
  const setMcpServers = useSettingsStore((s) => s.setMcpServers)

  const [models, setModels] = useState<ModelInfo[]>([])
  const [modelsLoading, setModelsLoading] = useState(false)
  const [modelsError, setModelsError] = useState<string | null>(null)

  // Live-fetch the model list when the relevant inputs change.
  useEffect(() => {
    let cancelled = false
    if (!FETCHES_LIVE_MODELS[providerId]) {
      setModels(providerId === 'anthropic' ? ANTHROPIC_MODELS : [])
      setModelsLoading(false)
      setModelsError(null)
      return
    }
    if (!apiKey) {
      setModels([])
      setModelsLoading(false)
      setModelsError(null)
      return
    }
    setModelsLoading(true)
    setModelsError(null)
    listModelsForProvider({
      providerId,
      apiKey,
      ...(baseUrl ? { baseUrl } : {}),
    })
      .then((ms) => {
        if (cancelled) return
        setModels(ms)
        setModelsLoading(false)
        // If the currently-selected model isn't in the list, fall
        // back to the provider's default.
        if (ms.length > 0 && !ms.some((m) => m.id === model)) {
          setModel(DEFAULT_MODEL[providerId] || ms[0]!.id)
        }
      })
      .catch((err: unknown) => {
        if (cancelled) return
        setModels([])
        setModelsLoading(false)
        setModelsError(err instanceof Error ? err.message : String(err))
      })
    return () => {
      cancelled = true
    }
  }, [providerId, apiKey, baseUrl, model, setModel])

  const onProviderChange = (next: ProviderId): void => {
    setProviderId(next)
    // Switch the model to the new provider's default so the
    // dropdown stays coherent.
    setModel(DEFAULT_MODEL[next] || '')
  }

  return (
    <div className="h-full overflow-y-auto p-4 space-y-4 text-sm">
      <h2 className="text-lg font-semibold">Settings</h2>

      <div>
        <label className="block font-medium" htmlFor="settings-provider">
          Provider
        </label>
        <select
          id="settings-provider"
          value={providerId}
          onChange={(e) => onProviderChange(e.target.value as ProviderId)}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1"
        >
          <option value="anthropic">Anthropic</option>
          <option value="minimax">MiniMax (M3)</option>
          <option value="openai">OpenAI</option>
          <option value="openai-compatible">OpenAI-compatible</option>
        </select>
      </div>

      <div>
        <label className="block font-medium" htmlFor="settings-apikey">
          API key
        </label>
        <input
          id="settings-apikey"
          type="password"
          value={apiKey}
          onChange={(e) => setApiKey(e.target.value)}
          placeholder="sk-ant-… or eyJ…"
          autoComplete="off"
          spellCheck={false}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1 font-mono"
        />
        <p className="mt-1 text-xs text-neutral-500">
          Stored in localStorage on this device. Not synced.
        </p>
      </div>

      {HAS_BASE_URL[providerId] ? (
        <div>
          <label className="block font-medium" htmlFor="settings-baseurl">
            Base URL
          </label>
          <input
            id="settings-baseurl"
            value={baseUrl}
            onChange={(e) => setBaseUrl(e.target.value)}
            placeholder="https://api.minimaxi.com/v1"
            autoComplete="off"
            spellCheck={false}
            className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1 font-mono"
          />
          <p className="mt-1 text-xs text-neutral-500">
            Leave empty to use the provider default.
          </p>
        </div>
      ) : null}

      <div>
        <label className="block font-medium" htmlFor="settings-model">
          Model
        </label>
        <select
          id="settings-model"
          value={model}
          onChange={(e) => setModel(e.target.value)}
          disabled={modelsLoading || models.length === 0}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1 disabled:bg-neutral-50 disabled:text-neutral-400"
        >
          {modelsLoading ? (
            <option>Loading…</option>
          ) : models.length === 0 ? (
            <option value={model}>{model || '(add API key to load models)'}</option>
          ) : (
            models.map((m) => (
              <option key={m.id} value={m.id}>
                {m.displayName}
                {m.contextWindow ? ` (${formatCtx(m.contextWindow)})` : ''}
              </option>
            ))
          )}
        </select>
        {modelsError ? (
          <p className="mt-1 text-xs text-red-600">
            Could not load model list: {modelsError}
          </p>
        ) : null}
      </div>

      <hr className="border-neutral-200" />

      <div>
        <div className="flex items-center justify-between">
          <label className="block font-medium">MCP servers</label>
          <button
            type="button"
            onClick={() => {
              setMcpServers([
                ...mcpServers,
                { name: `server-${mcpServers.length + 1}`, url: '', transport: 'http' },
              ])
            }}
            className="rounded bg-neutral-100 px-2 py-0.5 text-xs hover:bg-neutral-200"
          >
            + Add
          </button>
        </div>
        {mcpServers.length === 0 ? (
          <p className="mt-1 text-xs text-neutral-500">
            None configured. The LLM only sees built-in office.js
            skills.
          </p>
        ) : (
          <ul className="mt-2 space-y-2">
            {mcpServers.map((s, i) => (
              <li
                key={`${s.name}-${i}`}
                className="rounded border border-neutral-200 p-2"
              >
                <div className="flex gap-1">
                  <input
                    value={s.name}
                    onChange={(e) => {
                      const next = [...mcpServers]
                      next[i] = { ...next[i]!, name: e.target.value }
                      setMcpServers(next)
                    }}
                    placeholder="name"
                    className="flex-1 rounded border border-neutral-200 px-1.5 py-0.5 text-xs"
                  />
                  <button
                    type="button"
                    onClick={() => {
                      setMcpServers(mcpServers.filter((_, j) => j !== i))
                    }}
                    className="rounded text-xs text-red-600 hover:bg-red-50 px-1.5"
                  >
                    ×
                  </button>
                </div>
                <input
                  value={s.url}
                  onChange={(e) => {
                    const next = [...mcpServers]
                    next[i] = { ...next[i]!, url: e.target.value }
                    setMcpServers(next)
                  }}
                  placeholder="https://…/mcp"
                  className="mt-1 block w-full rounded border border-neutral-200 px-1.5 py-0.5 font-mono text-xs"
                />
                <select
                  value={s.transport}
                  onChange={(e) => {
                    const next = [...mcpServers]
                    next[i] = {
                      ...next[i]!,
                      transport: e.target.value as 'http' | 'sse',
                    }
                    setMcpServers(next)
                  }}
                  className="mt-1 block w-full rounded border border-neutral-200 px-1.5 py-0.5 text-xs"
                >
                  <option value="http">HTTP</option>
                  <option value="sse">SSE</option>
                </select>
              </li>
            ))}
          </ul>
        )}
        <p className="mt-1 text-xs text-neutral-500">
          Each tool exposed by these MCP servers is exposed to the
          LLM alongside the built-in office.js skills. Spec §13 W27.
        </p>
      </div>
    </div>
  )
}

function formatCtx(tokens: number): string {
  if (tokens >= 1_000_000) return `${(tokens / 1_000_000).toFixed(1)}M ctx`
  if (tokens >= 1_000) return `${Math.round(tokens / 1_000)}K ctx`
  return `${tokens} ctx`
}
