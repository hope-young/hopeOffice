/**
 * Skill registry. Source of truth: SPEC §7.
 *
 * The orchestrator pulls the skills relevant to the current host via
 * `listSkillsForHost()` and converts them to AI SDK 5 tool definitions
 * (see `core/agent/tools.ts`).
 *
 * Each skill is a plain TypeScript function with a Zod args schema
 * and an explicit `host` allow-list. The same name is used in the
 * registry key and on the tool surface so the LLM sees
 * `host-skill` style identifiers that match the host context.
 */
import type { HostKind } from '../types'
import type { Skill } from './types'

import { addChart } from './excel/add-chart'
import { addTable as excelAddTable } from './excel/add-table'
import { addFormula } from './excel/add-formula'
import { setCellValue } from './excel/set-cell-value'
import { sortRange } from './excel/sort-range'
import { findText as excelFindText } from './excel/find-text'

import { addText as pptAddText } from './powerpoint/add-text'
import { addSlide } from './powerpoint/add-slide'
import { duplicateSlide } from './powerpoint/duplicate-slide'
import { pptAddTable } from './powerpoint/pp-add-table'
import { pptAddBullets } from './powerpoint/pp-add-bullets'
import { pptAddImage } from './powerpoint/pp-add-image'

import { addTable as wordAddTable } from './word/add-table'
import { addText as wordAddText } from './word/add-text'
import { replaceText } from './word/replace-text'
import { addPageBreak } from './word/add-page-break'
import { findText as wordFindText } from './word/find-text'
import { setFont } from './word/wd-set-font'
import { addImage as wordAddImage } from './word/wd-add-image'

export const SKILL_REGISTRY = {
  // --- Excel ---
  'add-chart': addChart,
  'xl-add-table': excelAddTable,
  'add-formula': addFormula,
  'set-cell-value': setCellValue,
  'sort-range': sortRange,
  'xl-find-text': excelFindText,
  // --- PowerPoint ---
  'ppt-add-text': pptAddText,
  'add-slide': addSlide,
  'duplicate-slide': duplicateSlide,
  'ppt-add-table': pptAddTable,
  'ppt-add-bullets': pptAddBullets,
  'ppt-add-image': pptAddImage,
  // --- Word ---
  'word-add-table': wordAddTable,
  'word-add-text': wordAddText,
  'replace-text': replaceText,
  'add-page-break': addPageBreak,
  'word-find-text': wordFindText,
  'set-font': setFont,
  'add-image': wordAddImage,
} as const

export type SkillName = keyof typeof SKILL_REGISTRY

/**
 * Skills the LLM is allowed to call in the current host. The LLM
 * only ever sees this filtered list — a Word skill never appears
 * in an Excel conversation, etc.
 */
export function listSkillsForHost(host: HostKind): Skill<unknown, unknown>[] {
  const all = Object.values(SKILL_REGISTRY) as Skill<unknown, unknown>[]
  return all.filter((s) => s.host.includes(host))
}
