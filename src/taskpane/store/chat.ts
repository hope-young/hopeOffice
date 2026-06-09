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
 *
 * Phase 6: history persistence. Committed messages (post
 * `stream-end`) are mirrored to localStorage so the user can
 * refresh or close + reopen the task pane without losing the
 * conversation. Capped at MAX_HISTORY to keep storage small.
 */
import { create } from 'zustand'
import { Orchestrator } from '@core/agent/orchestrator'
import { createProvider } from '@core/providers/registry'
import type { AgentState, HostKind, Message } from '@core/types'
import { useSettingsStore } from './settings'

const HISTORY_KEY = 'hope-office_history'
const MAX_HISTORY = 100
const SAVE_DEBOUNCE_MS = 500

// ---------- History (localStorage) ----------

function loadHistory(): Message[] | null {
  if (typeof localStorage === 'undefined') return null
  try {
    const raw = localStorage.getItem(HISTORY_KEY)
    if (!raw) return null
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return null
    // Cast is safe-ish: we only round-trip Message objects we
    // wrote ourselves, but stale-shape rows could exist if the
    // schema changed across versions. A lightweight check keeps
    // the reducer from choking on bad data.
    return parsed.filter((m): m is Message =>
      m != null &&
      typeof m === 'object' &&
      'role' in m && 'content' in m &&
      typeof (m as { role: unknown }).role === 'string' &&
      typeof (m as { content: unknown }).content === 'string'
    ).slice(-MAX_HISTORY)
  } catch {
    return null
  }
}

function saveHistory(messages: Message[]): void {
  if (typeof localStorage === 'undefined') return
  try {
    const trimmed = messages.slice(-MAX_HISTORY)
    localStorage.setItem(HISTORY_KEY, JSON.stringify(trimmed))
  } catch {
    // Quota exceeded or storage disabled — silently no-op; in
    // production we'd want to surface this somewhere.
  }
}

function clearHistory(): void {
  if (typeof localStorage === 'undefined') return
  try {
    localStorage.removeItem(HISTORY_KEY)
  } catch {
    // ignore
  }
}

// ---------- Orchestrator singleton (module scope) ----------

/**
 * The host is set once at Office.onReady time (see
 * `setOrchestratorHost` below). We keep it in module scope so the
 * orchestrator can read it from inside `streamChat` without the
 * Zustand store needing to know about Office.
 */
let currentHost: HostKind = 'unsupported'

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
  getHost: () => currentHost,
  getMcpServers: () => useSettingsStore.getState().mcpServers,
})

/**
 * Called by `src/taskpane/index.tsx` once Office.onReady fires.
 * Dev-mode browser preview keeps the default `unsupported`.
 */
export function setOrchestratorHost(host: HostKind): void {
  currentHost = host
}

/**
 * Called by ChatPanel's Reset button (or anywhere else that wants
 * to wipe the conversation + the persisted history). Replaces the
 * in-memory messages with `[]` and clears localStorage.
 */
export function clearChatHistory(): void {
  clearHistory()
  orchestrator.dispatch({ type: 'restore-messages', messages: [] })
}

// Restore prior history on module load. We dispatch after the
// store + orchestrator are wired so the resulting state lands in
// the same pipeline as a normal user action.
const restored = loadHistory()
if (restored && restored.length > 0) {
  // queueMicrotask: avoid dispatching during module init, when
  // the orchestrator subscriber chain may not be fully attached.
  queueMicrotask(() => {
    orchestrator.dispatch({ type: 'restore-messages', messages: restored })
  })
}

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
//
// We also debounce-save committed messages to localStorage so the
// user can refresh the task pane and pick up where they left off.
let saveTimer: ReturnType<typeof setTimeout> | null = null
orchestrator.subscribe((state) => {
  useChatStore.setState({ state })
  if (saveTimer) clearTimeout(saveTimer)
  saveTimer = setTimeout(() => {
    saveHistory(state.messages)
  }, SAVE_DEBOUNCE_MS)
})

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
