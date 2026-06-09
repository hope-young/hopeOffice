/**
 * Orchestrator — the side-effect half of the agent. Source of
 * truth: SPEC §5.
 *
 * The reducer is a pure function over AgentState; this class is the
 * one place where state actually changes by talking to the outside
 * world. Specifically:
 *   - dispatches a 'user-send' event
 *   - opens a ChatProvider stream
 *   - translates every StreamChunk into an AgentEvent
 *   - dispatches the events one by one (reducer applies them in order)
 *   - emits the final 'stream-end' / 'error'
 *
 * What this Phase-4 version does NOT do:
 *   - Forward `tools` to the provider. The orchestrator asks the
 *     provider to stream text only; skill execution is Phase 5.
 *   - Run Skills in the sandbox. `sandboxExecutor.runOnHost` is still
 *     a stub; the orchestrator would have to dispatch tool-call
 *     results back into a fresh LLM turn, which is Phase 5's work.
 *   - Use `stepCountIs` for multi-step tool loops. Single-turn only.
 *
 * Subscriber model: a single listener (the Zustand chat store) gets
 * the new state after every dispatched event. Synchronous dispatch
 * keeps React renders coherent; the async loop yields to the event
 * loop between chunks.
 */
import { INITIAL_AGENT_STATE, reduce } from './reducer'
import type { ChatProvider, StreamChatOpts } from '../providers/interface'
import type { AgentEvent, AgentState, HostKind } from '../types'
import { toolsForHost } from './tools'
import { buildSystemPrompt } from './system-prompt'

export type OrchestratorDeps = {
  /** Provider factory: returns the current ChatProvider or null when
   *  the user hasn't configured one in Settings. Re-evaluated on
   *  every send so settings changes take effect without restart. */
  getProvider: () => ChatProvider | null
  /** Provider-scoped model id from Settings. */
  getModel: () => string
  /** Current Office host. Set once at Office.onReady; falls back
   *  to 'unsupported' in dev-mode browser preview. */
  getHost: () => HostKind
}

export class Orchestrator {
  private state: AgentState
  private listener: ((state: AgentState) => void) | null = null
  private ac: AbortController | null = null

  constructor(private deps: OrchestratorDeps) {
    this.state = { ...INITIAL_AGENT_STATE }
  }

  getState(): AgentState {
    return this.state
  }

  subscribe(fn: (state: AgentState) => void): () => void {
    this.listener = fn
    fn(this.state)
    return () => {
      if (this.listener === fn) this.listener = null
    }
  }

  /**
   * Public so the chat store can dispatch `restore-messages` /
   * `clear` events from outside the orchestrator's normal
   * user-send flow. Not part of the intended API surface for
   * UI components — use `send` / `abort` / `reset`.
   */
  dispatch(event: AgentEvent): void {
    this.state = reduce(this.state, event)
    this.listener?.(this.state)
  }

  /**
   * Send a user turn. No-op if the orchestrator is mid-turn (status
   * not 'idle'). The previous turn is auto-aborted if you really
   * insist — but in practice the UI button should be disabled in
   * that case.
   */
  async send(text: string): Promise<void> {
    if (this.state.status !== 'idle' && this.state.status !== 'error') return

    this.ac?.abort()
    this.ac = new AbortController()
    const signal = this.ac.signal

    this.dispatch({ type: 'user-send', text })

    const provider = this.deps.getProvider()
    if (!provider) {
      this.dispatch({
        type: 'error',
        error: new Error(
          'No provider configured. Open Settings to set an API key.',
        ),
      })
      return
    }

    try {
      const host = this.deps.getHost()
      const opts: StreamChatOpts = {
        messages: this.state.messages,
        tools: toolsForHost(host),
        system: { content: buildSystemPrompt(host) },
        model: this.deps.getModel(),
        signal,
        maxSteps: 5,
      }
      const stream = provider.streamChat(opts)

      for await (const chunk of stream) {
        if (signal.aborted) return
        switch (chunk.type) {
          case 'text-delta':
            this.dispatch({ type: 'stream-token', token: chunk.delta })
            break
          case 'reasoning-delta':
            // Reasoning models (e.g. MiniMax M3) emit a separate
            // delta stream for `<think>`-style blocks. The reducer
            // stores it on `draftReasoning` and the ChatPanel
            // renders it in a collapsible <details> block. Don't
            // fold it into `state.draft` — that would mix
            // reasoning into the visible answer.
            this.dispatch({ type: 'reasoning-delta', delta: chunk.delta })
            break
          case 'tool-call-start':
            this.dispatch({ type: 'tool-call-start', toolCall: chunk.toolCall })
            break
          case 'tool-call-args':
            // Phase 4: the provider doesn't emit args deltas (we
            // never advertise tools). Ignore for now.
            break
          case 'tool-call-result':
            this.dispatch({
              type: 'tool-call-result',
              toolCallId: chunk.toolCallId,
              result: chunk.result,
            })
            break
          case 'tool-call-error':
            this.dispatch({
              type: 'tool-call-error',
              toolCallId: chunk.toolCallId,
              error: chunk.error,
            })
            break
          case 'usage':
            // Phase 4: cost is a no-op. Phase 6 accumulates usage
            // into agent cost via a dedicated reducer path.
            break
          case 'finish':
            this.dispatch({ type: 'stream-end' })
            if (chunk.reason === 'error') {
              this.dispatch({
                type: 'error',
                error: new Error(chunk.error ?? 'stream finished with error'),
              })
            }
            return
        }
      }

      // Stream completed without an explicit 'finish' chunk — end
      // the turn so the draft gets committed.
      this.dispatch({ type: 'stream-end' })
    } catch (err) {
      this.dispatch({
        type: 'error',
        error: err instanceof Error ? err : new Error(String(err)),
      })
    }
  }

  /** Abort the in-flight turn, if any. */
  abort(): void {
    this.ac?.abort()
  }

  /** Wipe the draft / toolCalls / error. Messages survive. */
  reset(): void {
    this.ac?.abort()
    this.dispatch({ type: 'reset' })
  }
}
