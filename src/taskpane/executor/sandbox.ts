/**
 * Parent end of the postMessage bridge to the sandbox iframe.
 *
 * Source of truth: SPEC_DETAILS.md §2.
 *
 * Mounts `src/taskpane/executor/iframe.html` into the DOM, waits for
 * the child to send `{type: 'ready'}`, then exposes `runInSandbox`
 * for one-shot code execution.
 *
 * What this module does NOT do (deferred to later phases):
 *   - Serialize a Skill's `execute(args, ctx)` function into a string
 *     the child can `new Function(...)` over. The TS function
 *     references its closure (Zod-parsed args, `ctx`, etc.) and can't
 *     be naively toString'd. Phase 4 (orchestrator) will own the
 *     serialisation, Phase 5 (skill registry) the per-skill format.
 *   - Implement `sandboxExecutor.runOnHost` for real. A stub is
 *     exported so downstream callers can wire it up; calling it
 *     today throws.
 */
import type { HostKind } from '../../core/types'
import type { Executor, OfficeHostAPI } from '../../core/skills/types'

const PENDING_TIMEOUT_MS = 30_000
const READY_TIMEOUT_MS = 30_000
const IFRAME_SRC = '/src/taskpane/executor/iframe.html'

// ---------- Protocol types ----------

export type ParentToIframe =
  | { type: 'execute'; id: string; code: string; host: HostKind }
  | { type: 'ping'; id: string }

export type IframeToParent =
  | { type: 'ready' }
  | { type: 'pong'; id: string }
  | { type: 'host'; host: HostKind | 'unsupported' }
  | { type: 'result'; id: string; ok: true; value: unknown }
  | {
      type: 'result'
      id: string
      ok: false
      error: { message: string; stack?: string }
    }

// ---------- Module state (singleton per taskpane session) ----------

let iframeRef: HTMLIFrameElement | null = null
let readyPromise: Promise<void> | null = null

function ensureMounted(): void {
  if (typeof window === 'undefined') {
    throw new Error('runInSandbox: only callable in a browser context')
  }
  if (iframeRef) return

  const iframe = document.createElement('iframe')
  // allow-scripts: the child runs user code.
  // allow-same-origin: needed so Office.js (loaded by the child) sees
  //   the real localhost origin, not a null origin that would defeat
  //   its own init logic.
  iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin')
  iframe.setAttribute('aria-hidden', 'true')
  iframe.title = 'hope-Office sandbox'
  // Visually hide but keep the element in the document tree.
  iframe.style.cssText =
    'position:fixed;bottom:0;right:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none;'
  iframe.src = IFRAME_SRC
  document.body.appendChild(iframe)
  iframeRef = iframe

  readyPromise = new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => {
      window.removeEventListener('message', onMessage)
      reject(new Error(`Sandbox iframe never sent "ready" within ${READY_TIMEOUT_MS / 1000}s`))
    }, READY_TIMEOUT_MS)

    function onMessage(e: MessageEvent) {
      const msg = e.data as IframeToParent
      if (msg && msg.type === 'ready') {
        window.removeEventListener('message', onMessage)
        clearTimeout(timer)
        resolve()
      }
    }
    window.addEventListener('message', onMessage)
  })
}

// ---------- Public API ----------

/**
 * Run an arbitrary string of code inside the sandbox iframe, scoped
 * to the given host. The child wraps the code in the host's
 * `Word.run` / `Excel.run` / `PowerPoint.run` (or evals it raw when
 * Office.js isn't available, e.g. dev-mode preview).
 *
 * Rejects on:
 *   - iframe never ready (READY_TIMEOUT_MS)
 *   - host mismatch between manifest and Office
 *   - user code throws or returns a rejected promise
 *   - call exceeds PENDING_TIMEOUT_MS
 *   - signal is aborted
 */
export async function runInSandbox(
  code: string,
  host: HostKind,
  signal: AbortSignal,
): Promise<unknown> {
  ensureMounted()
  await readyPromise

  const id = crypto.randomUUID()

  return new Promise<unknown>((resolve, reject) => {
    let settled = false

    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      window.removeEventListener('message', listener)
      signal.removeEventListener('abort', onAbort)
      reject(new Error(`Sandbox call timed out after ${PENDING_TIMEOUT_MS / 1000}s`))
    }, PENDING_TIMEOUT_MS)

    function onAbort() {
      if (settled) return
      settled = true
      clearTimeout(timer)
      window.removeEventListener('message', listener)
      reject(new Error('Aborted'))
    }
    signal.addEventListener('abort', onAbort, { once: true })

    function listener(e: MessageEvent) {
      const msg = e.data as IframeToParent
      if (!msg || msg.type !== 'result' || msg.id !== id) return
      if (settled) return
      settled = true
      clearTimeout(timer)
      window.removeEventListener('message', listener)
      signal.removeEventListener('abort', onAbort)

      if (msg.ok) {
        resolve(msg.value)
      } else {
        const err = new Error(msg.error.message)
        if (msg.error.stack) err.stack = msg.error.stack
        reject(err)
      }
    }
    window.addEventListener('message', listener)

    const target = iframeRef!.contentWindow
    if (!target) {
      settled = true
      clearTimeout(timer)
      window.removeEventListener('message', listener)
      signal.removeEventListener('abort', onAbort)
      reject(new Error('Sandbox iframe has no contentWindow (detached?)'))
      return
    }
    const message: ParentToIframe = { type: 'execute', id, code, host }
    target.postMessage(message, '*')
  })
}

/**
 * `Executor` implementation backed by `runInSandbox`. The `runOnHost`
 * body is a stub for now: serialising a live TS function (with its
 * closure over `args` and `ctx`) into a string the child iframe can
 * `new Function(...)` is non-trivial and is deferred to Phase 4
 * (orchestrator) + Phase 5 (skill registry). The stub keeps the
 * `Executor` interface type-checkable so skill code can already
 * import it without TypeScript complaints.
 */
export const sandboxExecutor: Executor = {
  async runOnHost<T>(
    _host: HostKind,
    _fn: (hostApi: OfficeHostAPI) => Promise<T>,
  ): Promise<T> {
    throw new Error('sandboxExecutor.runOnHost is not yet implemented (Phase 4)')
  },
}
