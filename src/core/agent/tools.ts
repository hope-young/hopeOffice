/**
 * Skill → AI SDK 5 tool conversion. Source of truth: SPEC §6 +
 * SPEC_DETAILS §6.
 *
 * Each skill is exposed to the LLM as its own tool (one tool per
 * skill name) — simpler than a single `execute_skill(name, args)`
 * wrapper, and the LLM picks skills more reliably when they have
 * their own typed signature.
 *
 * The AI SDK 5 `Tool.execute` function runs locally in the task
 * pane (NOT in the sandbox iframe). That's deliberate for Phase 5:
 *  - office.js only works inside the Office task pane anyway, so
 *    a sandbox detour would just add round-trip latency.
 *  - The Function-serialisation problem (`fn.toString()` loses
 *    closures over skill args + ctx) becomes a non-issue — the
 *    AI SDK passes the LLM's tool-call input through the tool's
 *    `inputSchema` (Zod) into our `execute(input, ctx)` directly.
 *  - Phase 6+ can swap to a sandboxed executor for tighter
 *    isolation by changing only this file.
 */
import { tool, type Tool } from 'ai'
import type { HostKind } from '../types'
import type { Skill } from '../skills/types'
import { listSkillsForHost } from '../skills'

/** Stub executor — Phase 5 skills don't go through it. */
const STUB_EXECUTOR = {
  async runOnHost(): Promise<never> {
    throw new Error(
      'Stub executor: skills in Phase 5 call office.js directly. ' +
        'A real sandbox executor lands in Phase 6+.',
    )
  },
} as const

/**
 * Build the tools map for the current host. Pass it straight to
 * `streamText({ tools: toolsForHost(host) })`.
 */
export function toolsForHost(host: HostKind): Record<string, Tool> {
  const out: Record<string, Tool> = {}
  for (const s of listSkillsForHost(host)) {
    out[s.name] = skillToTool(s)
  }
  return out
}

function skillToTool(s: Skill<unknown, unknown>): Tool {
  // AI SDK 5's Tool<INPUT, OUTPUT> generic lets us type the input
  // from the skill's Zod schema (INPUT) and the skill's return type
  // (OUTPUT). `tool()` is an identity wrapper that just preserves
  // the type at the call site.
  return tool({
    description: s.description,
    inputSchema: s.args,
    execute: async (input: unknown) => {
      // The AI SDK has already parsed + validated `input` against
      // `s.args` (Zod). We pass the typed value straight through.
      return await s.execute(input as never, {
        executor: STUB_EXECUTOR,
        conversationId: 'pending', // Phase 8 wires real conversation ids
        signal: new AbortController().signal, // not abortable per-call yet
      })
    },
  })
}
