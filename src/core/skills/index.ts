/**
 * Skill registry. Source of truth: SPEC §7.
 *
 * The orchestrator pulls the skills relevant to the current host via
 * `listSkillsForHost()` and converts them to AI SDK 5 tool definitions
 * (see `core/agent/tools.ts`). Each skill is a plain TypeScript
 * function with a Zod args schema and an explicit `host` allow-list.
 *
 * Phase 5: Excel only (per the user's priority order — Word + PPT
 * land in later phases). The first skill is `add-chart`, validating
 * the LLM→skill→office.js call chain end-to-end.
 */
import type { HostKind } from '../types'
import type { Skill } from './types'
import { addChart } from './excel/add-chart'
import { addText as addTextPPT } from './powerpoint/add-text'
import { addSlide } from './powerpoint/add-slide'
import { addTable } from './word/add-table'
import { addText as addTextWord } from './word/add-text'

export const SKILL_REGISTRY = {
  'add-chart': addChart,
  // PowerPoint and Word both expose an "add-text" skill, but
  // they live on different hosts and have different arg shapes
  // (font size, position semantics). The skill name itself
  // disambiguates with a host prefix.
  'ppt-add-text': addTextPPT,
  'add-slide': addSlide,
  'add-table': addTable,
  'word-add-text': addTextWord,
} as const

export type SkillName = keyof typeof SKILL_REGISTRY

/**
 * Skills the LLM is allowed to call in the current host. The LLM
 * only ever sees this filtered list — a Word skill never appears
 * in an Excel conversation.
 */
export function listSkillsForHost(host: HostKind): Skill<unknown, unknown>[] {
  const all = Object.values(SKILL_REGISTRY) as Skill<unknown, unknown>[]
  return all.filter((s) => s.host.includes(host))
}
