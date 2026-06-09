/**
 * Tests for the system-prompt builder. We just verify that the
 * right skills are listed for the right host — the exact wording
 * is intentionally not pinned because users will want to tweak
 * it without breaking these tests.
 */
import { describe, expect, it } from 'vitest'
import { buildSystemPrompt } from './system-prompt'

describe('buildSystemPrompt', () => {
  it('lists Excel skills when host=excel', () => {
    const p = buildSystemPrompt('excel')
    expect(p).toContain('add-chart')
    expect(p).not.toContain('add-text')
    expect(p).not.toContain('add-slide')
    expect(p.toLowerCase()).toContain('excel')
  })

  it('lists PowerPoint skills when host=powerpoint', () => {
    const p = buildSystemPrompt('powerpoint')
    expect(p).toContain('add-text')
    expect(p).toContain('add-slide')
    expect(p).not.toContain('add-chart')
    expect(p.toLowerCase()).toContain('powerpoint')
  })

  it('marks empty when no skills are registered for the host', () => {
    const p = buildSystemPrompt('word')
    expect(p.toLowerCase()).toContain('no skills')
  })

  it('inlines the Zod-derived JSON schema for each skill', () => {
    const p = buildSystemPrompt('excel')
    // z.toJSONSchema emits a JSON Schema with `properties` — we
    // check for that as a proxy for "the schema is inlined".
    expect(p).toMatch(/"properties"/)
    expect(p).toMatch(/"sheetName"/)
    expect(p).toMatch(/"dataRange"/)
  })
})
