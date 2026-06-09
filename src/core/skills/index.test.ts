/**
 * Tests for the skill registry + host filter. These don't touch
 * office.js (jsdom doesn't have it) — they just exercise the
 * pure logic in core/skills/index.ts.
 */
import { describe, expect, it } from 'vitest'
import { SKILL_REGISTRY, listSkillsForHost } from './index'

describe('SKILL_REGISTRY', () => {
  it('contains the five Phase 5 / 5b / 5c skills', () => {
    expect(Object.keys(SKILL_REGISTRY).sort()).toEqual(
      [
        'add-chart',
        'add-slide',
        'add-table',
        'ppt-add-text',
        'word-add-text',
      ].sort(),
    )
  })

  it('every skill declares a non-empty name + description + host list', () => {
    for (const [name, skill] of Object.entries(SKILL_REGISTRY)) {
      expect(skill.name, `${name}.name`).toBe(name)
      expect(skill.description.length, `${name}.description`).toBeGreaterThan(0)
      expect(skill.host.length, `${name}.host`).toBeGreaterThan(0)
      // Spread the readonly array so vitest's mutable arrayContaining
      // signature accepts it.
      expect(['word', 'excel', 'powerpoint']).toEqual(
        expect.arrayContaining([...skill.host]),
      )
    }
  })
})

describe('listSkillsForHost', () => {
  it('returns only Excel skills for the excel host', () => {
    const skills = listSkillsForHost('excel')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual(['add-chart'])
  })

  it('returns PPT skills for the powerpoint host (host-prefixed key)', () => {
    const skills = listSkillsForHost('powerpoint')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual(['add-slide', 'ppt-add-text'])
  })

  it('returns Word skills for the word host (host-prefixed key)', () => {
    const skills = listSkillsForHost('word')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual(['add-table', 'word-add-text'])
  })

  it('returns nothing for unsupported', () => {
    expect(listSkillsForHost('unsupported')).toEqual([])
  })

  it('the word host does NOT include the PPT add-text skill', () => {
    const skills = listSkillsForHost('word')
    const names = skills.map((s) => s.name)
    expect(names).not.toContain('ppt-add-text')
  })
})
