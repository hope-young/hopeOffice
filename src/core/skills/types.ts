/**
 * Skill — the typed, host-filtered operation the LLM can invoke.
 *
 * Source of truth: SPEC §7 + SPEC_DETAILS §6.
 * Skills are plain TypeScript functions; the LLM calls them through
 * the `execute_skill` tool with args that Zod parses before the body
 * runs (so a malformed tool call is rejected before any office.js
 * code executes).
 *
 * No DOM, no React imports allowed in this file. The body itself may
 * import `@microsoft/office-js` since `host` is a `HostKind`.
 */
import type { z } from 'zod'
import type { HostKind } from '../types'

// ---------- Executor (forward-declared; full impl in Phase 3) ----------

/**
 * Tagged union of the host Office.js namespaces a skill can ask for.
 * `unsupported` is included so dev-mode runners (browser, tests) can
 * type-check without forcing a real Office binding.
 */
export type OfficeHostAPI =
  | { kind: 'word'; ns: typeof Word }
  | { kind: 'excel'; ns: typeof Excel }
  | { kind: 'powerpoint'; ns: typeof PowerPoint }
  | { kind: 'unsupported'; ns: undefined }

/**
 * The execution abstraction skills depend on. In dev this runs the
 * function in-process; in production it marshals the function into
 * the sandbox iframe via the postMessage protocol from
 * SPEC_DETAILS §2.
 */
export interface Executor {
  runOnHost<T>(
    host: HostKind,
    fn: (hostApi: OfficeHostAPI) => Promise<T>,
  ): Promise<T>
}

// ---------- Ctx (skill execution context) ----------

/**
 * Ctx — passed to every Skill.execute. Carries the runtime
 * dependencies a skill needs without forcing it to import concrete
 * taskpane modules.
 */
export type Ctx = {
  /** Office.js execution abstraction (see above). */
  executor: Executor
  /** Current conversation id, for logging and conversation-scoped state. */
  conversationId: string
  /** AbortSignal cancelled when the user hits Stop or the
   *  conversation is closed. Skills should forward this into any
   *  long-running office.js calls. */
  signal: AbortSignal
}

// ---------- Skill ----------

export interface Skill<TArgs, TResult> {
  /** Tool name passed to the LLM, e.g. `'add-table'`. */
  name: string
  /** Free-text description included in the system prompt's skill list. */
  description: string
  /** Which Office hosts this skill is available on. The orchestrator
   *  filters the registry by host before exposing tools to the LLM,
   *  so a Word skill never appears in an Excel conversation. */
  host: readonly HostKind[]
  /** Zod schema parsed before the body runs. The inferred type
   *  (`z.infer<typeof args>`) is what `execute` receives as
   *  its first argument. */
  args: z.ZodType<TArgs>
  execute(args: TArgs, ctx: Ctx): Promise<TResult>
}
