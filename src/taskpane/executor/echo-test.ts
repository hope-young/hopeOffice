/**
 * Smoke test for the sandbox postMessage bridge.
 *
 * Runs a tiny code body in the iframe via `runInSandbox` and returns
 * the round-trip result. The point isn't the computation — it's
 * confirming the parent/child message channel works end-to-end for
 * a given host.
 *
 * Call this from the dev console, a Settings debug button, or a
 * vitest suite. It deliberately doesn't go through the skill
 * registry (that's Phase 5) — it's the layer beneath.
 *
 * In dev (no Office.js), pass `host: 'unsupported'` to take the
 * eval branch; the iframe will then match the parent's claim and
 * the round-trip still exercises the protocol.
 */
import type { HostKind } from '../../core/types'
import { runInSandbox } from './sandbox'

/**
 * Code body sent to the sandbox. Lives as a module constant (rather
 * than inside the function) so it has no closure on test-wrapper
 * locals — exactly the shape real user code will have once the
 * skill serialiser lands in Phase 4.
 */
const ECHO_CODE = `
return {
  message: 'echo from sandbox',
  ts: Date.now(),
  host: 'word',
}
`

export type EchoTestResult =
  | { ok: true; value: unknown; host: HostKind; elapsedMs: number }
  | { ok: false; error: string; stack?: string; host: HostKind; elapsedMs: number }

/**
 * Run the echo probe against the sandbox and time the round-trip.
 */
export async function runEchoTest(host: HostKind = 'word'): Promise<EchoTestResult> {
  const ac = new AbortController()
  const t0 = performance.now()
  try {
    const value = await runInSandbox(ECHO_CODE, host, ac.signal)
    return { ok: true, value, host, elapsedMs: performance.now() - t0 }
  } catch (err) {
    const e = err as Error
    return {
      ok: false,
      error: e.message,
      stack: e.stack,
      host,
      elapsedMs: performance.now() - t0,
    }
  }
}
