/**
 * Tests for the skill registry + host filter. These don't touch
 * office.js (jsdom doesn't have it) — they just exercise the
 * pure logic in core/skills/index.ts.
 */
import { describe, expect, it } from 'vitest'
import { SKILL_REGISTRY, listSkillsForHost } from './index'

describe('SKILL_REGISTRY', () => {
  it('contains the three Phase 5 skills', () => {
    expect(Object.keys(SKILL_REGISTRY).sort()).toEqual(
      ['add-chart', 'add-slide', 'add-text'].sort(),
    )
  })

  it('every skill declares a non-empty name + description + host list', () => {
    for (const [name, skill] of Object.entries(SKILL_REGISTRY)) {
      expect(skill.name, `${name}.name`).toBe(name)
      expect(skill.description.length, `${name}.description`).toBeGreaterThan(0)
      expect(skill.host.length, `${name}.host`).toBeGreaterThan(0)
      expect(['word', 'excel', 'powerpoint']).toEqual(
        expect.arrayContaining(skill.host),
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

  it('returns only PowerPoint skills for the powerpoint host', () => {
    const skills = listSkillsForHost('powerpoint')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual(['add-slide', 'add-text'])
  })

  it('returns nothing for the word host in Phase 5', () => {
    const skills = listSkillsForHost('word')
    expect(skills).toEqual([])
  })

  it('returns nothing for unsupported', () => {
    expect(listSkillsForHost('unsupported')).toEqual([])
  })
})
