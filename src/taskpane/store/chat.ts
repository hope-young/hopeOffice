/**
 * Chat store — Zustand. Owns the live AgentState and exposes the
 * high-level actions (send / abort / reset) that drive the
 * Orchestrator. Source of truth: SPEC §9.
 *
 * Why a singleton Orchestrator + a Zustand store on top:
 *   - The Orchestrator already has a subscribe/notify model; a thin
 *     Zustand wrapper gives us React-friendly selectors and lets
 *     components read `useChatStore(s => s.state.messages)` without
 *     re-rendering on draft changes they don't care about.
 *   - The Orchestrator lives outside the store (module scope) so
 *     its lifetime isn't tied to React's render tree — the
 *     subscription is set up once at module load.
 */
import { create } from 'zustand'
import { Orchestrator } from '@core/agent/orchestrator'
import { createProvider } from '@core/providers/registry'
import type { AgentState } from '@core/types'
import { useSettingsStore } from './settings'

// ---------- Orchestrator singleton (module scope) ----------

const orchestrator = new Orchestrator({
  getProvider: () => {
    const { providerId, apiKey, model } = useSettingsStore.getState()
    if (!apiKey) return null
    try {
      return createProvider({ providerId, apiKey, model })
    } catch {
      return null
    }
  },
  getModel: () => useSettingsStore.getState().model,
})

// ---------- Store ----------

export type ChatState = {
  state: AgentState
  send: (text: string) => Promise<void>
  abort: () => void
  reset: () => void
}

export const useChatStore = create<ChatState>(() => ({
  state: orchestrator.getState(),
  send: (text) => orchestrator.send(text),
  abort: () => orchestrator.abort(),
  reset: () => orchestrator.reset(),
}))

// Wire the orchestrator's pub/sub into the store AFTER the store
// exists. The earlier version put this above the `create()` call and
// tripped a TDZ: the subscribe callback closes over useChatStore,
// which isn't initialised until line 18 — but the subscribe() on
// line 17 fires synchronously with the initial state, throwing
// "Cannot access 'useChatStore' before initialization".
orchestrator.subscribe((state) => {
  // Push every state change into the store. Components select the
  // slice they care about via the Zustand selector.
  useChatStore.setState({ state })
})
