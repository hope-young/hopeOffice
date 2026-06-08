/**
 * Agent state reducer — pure function from (state, event) to next
 * state. Source of truth: SPEC §5.
 *
 * Rules of the road:
 *   - reducer is the ONLY place state transitions live
 *   - side effects (network, sandbox, approval UI) live in
 *     orchestrator.ts; reducer just records what happened
 *   - unknown event types are a TS error (the `never` exhaustiveness
 *     check at the bottom), so adding a new event without handling
 *     it surfaces as a build failure
 *   - `approval` is deliberately NOT in AgentState — it's a side-
 *     effect handle (a Promise resolver) and doesn't belong in
 *     serialisable state. The orchestrator owns it locally.
 */
import type { AgentEvent, AgentState, Cost, Message, ToolCall } from '../types'

export const INITIAL_AGENT_STATE: AgentState = {
  messages: [],
  draft: '',
  toolCalls: [],
  status: 'idle',
  cost: emptyCost(),
}

function emptyCost(): Cost {
  return { input: 0, output: 0, total: 0, source: 'tokens-only' }
}

export function reduce(state: AgentState, event: AgentEvent): AgentState {
  switch (event.type) {
    case 'user-send': {
      // New user turn: append the message, clear the draft, switch
      // to streaming, drop any stale error.
      const userMsg: Message = { role: 'user', content: event.text }
      return {
        ...state,
        messages: [...state.messages, userMsg],
        draft: '',
        toolCalls: [],
        status: 'streaming',
        error: undefined,
      }
    }

    case 'stream-token': {
      // Ignore stragglers from a previous turn.
      if (state.status !== 'streaming') return state
      return { ...state, draft: state.draft + event.token }
    }

    case 'stream-end': {
      // Commit the draft as a finished assistant message. If the LLM
      // produced no text (e.g. it only emitted tool calls), don't add
      // an empty message — the tool-call-start events have already
      // updated state.toolCalls.
      if (!state.draft && state.toolCalls.length === 0) {
        return { ...state, status: 'idle' }
      }
      const assistantMsg: Message = state.toolCalls.length
        ? { role: 'assistant', content: state.draft, toolCalls: [...state.toolCalls] }
        : { role: 'assistant', content: state.draft }
      return {
        ...state,
        messages: [...state.messages, assistantMsg],
        draft: '',
        toolCalls: [],
        status: 'idle',
      }
    }

    case 'tool-call-start': {
      const tc: ToolCall = { ...event.toolCall, status: 'running', startedAt: Date.now() }
      // Replace if a placeholder already exists (keyed by id); else
      // append. This handles the case where the LLM sends a
      // tool-input-start first, then a tool-call.
      const existing = state.toolCalls.findIndex((t) => t.id === tc.id)
      const next =
        existing >= 0
          ? state.toolCalls.map((t, i) => (i === existing ? tc : t))
          : [...state.toolCalls, tc]
      return { ...state, toolCalls: next, status: 'executing' }
    }

    case 'tool-call-result': {
      return {
        ...state,
        toolCalls: state.toolCalls.map((t) =>
          t.id === event.toolCallId
            ? { ...t, result: event.result, status: 'success', completedAt: Date.now() }
            : t,
        ),
        // Return to streaming if the LLM is still emitting; orchestrator
        // is the source of truth for what comes next.
        status: state.status === 'executing' ? 'streaming' : state.status,
      }
    }

    case 'tool-call-error': {
      return {
        ...state,
        toolCalls: state.toolCalls.map((t) =>
          t.id === event.toolCallId
            ? {
                ...t,
                status: 'error',
                error: { message: event.error },
                completedAt: Date.now(),
              }
            : t,
        ),
        status: state.status === 'executing' ? 'streaming' : state.status,
      }
    }

    case 'approval-requested': {
      return { ...state, status: 'awaiting-approval' }
    }

    case 'approval-decided': {
      // Only meaningful when we were actually awaiting approval.
      if (state.status !== 'awaiting-approval') return state
      return { ...state, status: 'executing' }
    }

    case 'error': {
      return { ...state, status: 'error', error: event.error }
    }

    case 'reset': {
      return { ...INITIAL_AGENT_STATE, messages: state.messages }
    }
  }

  // Exhaustiveness check: if a new AgentEvent variant is added, TS
  // will fail the build until the switch above covers it.
  const _exhaustive: never = event
  return _exhaustive
}
