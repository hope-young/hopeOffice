/**
 * SettingsPanel — provider, API key, model. Persisted via the
 * Zustand settings store (localStorage). Phase 4 ships Anthropic
 * only; OpenAI / OpenAI-compatible slots are visible but disabled
 * with a hint about Phase 7.
 */
import type { ProviderId } from '@core/providers/registry'
import { useSettingsStore } from '../store/settings'

const ANTHROPIC_MODELS: Array<{ id: string; label: string }> = [
  { id: 'claude-opus-4-7', label: 'Claude Opus 4.7' },
  { id: 'claude-sonnet-4-5', label: 'Claude Sonnet 4.5' },
  { id: 'claude-haiku-4-5', label: 'Claude Haiku 4.5' },
]

export function SettingsPanel() {
  const providerId = useSettingsStore((s) => s.providerId)
  const apiKey = useSettingsStore((s) => s.apiKey)
  const model = useSettingsStore((s) => s.model)
  const baseUrl = useSettingsStore((s) => s.baseUrl)

  const setProviderId = useSettingsStore((s) => s.setProviderId)
  const setApiKey = useSettingsStore((s) => s.setApiKey)
  const setModel = useSettingsStore((s) => s.setModel)
  const setBaseUrl = useSettingsStore((s) => s.setBaseUrl)

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
          onChange={(e) => setProviderId(e.target.value as ProviderId)}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1"
        >
          <option value="anthropic">Anthropic</option>
          <option value="openai" disabled>
            OpenAI (Phase 7)
          </option>
          <option value="openai-compatible" disabled>
            OpenAI-compatible (Phase 7)
          </option>
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
          placeholder="sk-ant-…"
          autoComplete="off"
          spellCheck={false}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1 font-mono"
        />
        <p className="mt-1 text-xs text-neutral-500">
          Stored in localStorage on this device. Not synced.
        </p>
      </div>

      <div>
        <label className="block font-medium" htmlFor="settings-model">
          Model
        </label>
        <select
          id="settings-model"
          value={model}
          onChange={(e) => setModel(e.target.value)}
          className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1"
        >
          {ANTHROPIC_MODELS.map((m) => (
            <option key={m.id} value={m.id}>
              {m.label}
            </option>
          ))}
        </select>
      </div>

      {providerId === 'openai-compatible' ? (
        <div>
          <label className="block font-medium" htmlFor="settings-baseurl">
            Base URL
          </label>
          <input
            id="settings-baseurl"
            value={baseUrl}
            onChange={(e) => setBaseUrl(e.target.value)}
            placeholder="http://localhost:11434/v1"
            autoComplete="off"
            spellCheck={false}
            className="mt-1 block w-full rounded border border-neutral-300 px-2 py-1 font-mono"
          />
        </div>
      ) : null}
    </div>
  )
}
