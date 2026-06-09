/**
 * Settings store — Zustand + persist.
 *
 * Phase 9: persistence tier switches by environment.
 *   - dev / browser preview: `localStorage`
 *   - production (Office Add-in): `Office.context.roamingSettings`,
 *     which follows the user across machines.
 *
 * The store still exposes the same Zustand surface — we hide the
 * storage backend behind a small Storage adapter that mirrors
 * Zustand's `StateStorage` interface (getItem / setItem /
 * removeItem).
 */
import { create } from 'zustand'
import { createJSONStorage, persist, type StateStorage } from 'zustand/middleware'
import type { ProviderId } from '@core/providers/registry'
import type { McpServerConfig } from '@core/mcp/client'

export type SettingsState = {
  providerId: ProviderId
  apiKey: string
  /** Provider-scoped model id (e.g. 'claude-sonnet-4-5'). */
  model: string
  /** Only meaningful for openai-compatible. */
  baseUrl: string
  /** MCP server list (SPEC §13 W27). */
  mcpServers: McpServerConfig[]

  setProviderId: (id: ProviderId) => void
  setApiKey: (key: string) => void
  setModel: (m: string) => void
  setBaseUrl: (u: string) => void
  setMcpServers: (servers: McpServerConfig[]) => void
}

/**
 * Build the right `StateStorage` for the current environment.
 * Falls back to `localStorage` if Office is not present (dev /
 * browser preview) or if the Office storage is missing
 * (e.g. running the task pane outside a sideloaded context).
 */
export function makeStorage(): StateStorage {
  // Prefer Office's roaming settings when present. They follow
  // the user across machines (per Microsoft Learn) and survive
  // both browser refresh and full add-in restart.
  if (typeof Office !== 'undefined' && Office.context?.roamingSettings) {
    const settings = Office.context.roamingSettings
    return {
      getItem: (key: string): string | null => {
        const v = settings.get(key)
        return v == null ? null : String(v)
      },
      setItem: (key: string, value: string): void => {
        settings.set(key, value)
      },
      removeItem: (key: string): void => {
        settings.remove(key)
      },
    }
  }
  // Dev / browser fallback. StateStorage's signature wants a
  // Storage-like with getItem returning string | null, which
  // matches localStorage's interface (with a string | null
  // cast at the boundary).
  if (typeof localStorage !== 'undefined') {
    return {
      getItem: (key: string): string | null => localStorage.getItem(key),
      setItem: (key: string, value: string): void => {
        localStorage.setItem(key, value)
      },
      removeItem: (key: string): void => {
        localStorage.removeItem(key)
      },
    }
  }
  // No-op storage (e.g. SSR / some embedded WebView) — settings
  // simply won't persist. Better than throwing at import time.
  return {
    getItem: () => null,
    setItem: () => undefined,
    removeItem: () => undefined,
  }
}

const STORAGE_KEY = 'hope-office_settings'

export const useSettingsStore = create<SettingsState>()(
  persist(
    (set) => ({
      providerId: 'minimax',
      apiKey: '',
      model: 'MiniMax-M3',
      baseUrl: '',
      mcpServers: [],

      setProviderId: (id) => set({ providerId: id }),
      setApiKey: (key) => set({ apiKey: key }),
      setModel: (m) => set({ model: m }),
      setBaseUrl: (u) => set({ baseUrl: u }),
      setMcpServers: (servers) => set({ mcpServers: servers }),
    }),
    {
      name: STORAGE_KEY,
      storage: createJSONStorage(() => makeStorage()),
      partialize: (s) => ({
        providerId: s.providerId,
        apiKey: s.apiKey,
        model: s.model,
        baseUrl: s.baseUrl,
        mcpServers: s.mcpServers,
      }),
    },
  ),
)
