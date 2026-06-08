/**
 * Settings store — Zustand. Persists to localStorage.
 *
 * Source of truth: SPEC §9 + §11. We use localStorage uniformly
 * (dev + prod) for now; Phase 9 swaps to Office.context.roamingSettings
 * for production per SPEC §11. (Both work in the taskpane WebView
 * today; roamingSettings is the right home in a store-deploy add-in
 * because it follows the user across machines.)
 *
 * The store deliberately does NOT hold the ChatProvider instance
 * itself — it's reconstructed on demand by the chat store from
 * `providerId`/`apiKey`/`model` so settings edits take effect
 * without a page reload.
 */
import { create } from 'zustand'
import { createJSONStorage, persist } from 'zustand/middleware'
import type { ProviderId } from '@core/providers/registry'

export type SettingsState = {
  providerId: ProviderId
  apiKey: string
  /** Provider-scoped model id (e.g. 'claude-sonnet-4-5'). */
  model: string
  /** Only meaningful for openai-compatible. */
  baseUrl: string

  setProviderId: (id: ProviderId) => void
  setApiKey: (key: string) => void
  setModel: (m: string) => void
  setBaseUrl: (u: string) => void
}

const STORAGE_KEY = 'hope-office_settings'

export const useSettingsStore = create<SettingsState>()(
  persist(
    (set) => ({
      providerId: 'anthropic',
      apiKey: '',
      model: 'claude-sonnet-4-5',
      baseUrl: '',

      setProviderId: (id) => set({ providerId: id }),
      setApiKey: (key) => set({ apiKey: key }),
      setModel: (m) => set({ model: m }),
      setBaseUrl: (u) => set({ baseUrl: u }),
    }),
    {
      name: STORAGE_KEY,
      storage: createJSONStorage(() => localStorage),
      // Don't persist the setter functions (zustand does this by
      // default, but be explicit).
      partialize: (s) => ({
        providerId: s.providerId,
        apiKey: s.apiKey,
        model: s.model,
        baseUrl: s.baseUrl,
      }),
    },
  ),
)
