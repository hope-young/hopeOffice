/**
 * Tests for the skill registry + host filter.
 */
import { describe, expect, it } from 'vitest'
import { SKILL_REGISTRY, listSkillsForHost } from './index'

describe('SKILL_REGISTRY', () => {
  it('contains 19 skills (Phase 8)', () => {
    expect(Object.keys(SKILL_REGISTRY).length).toBe(19)
  })

  it('every skill declares a non-empty name + description + host list', () => {
    for (const [name, skill] of Object.entries(SKILL_REGISTRY)) {
      expect(skill.name, `${name}.name`).toBe(name)
      expect(skill.description.length, `${name}.description`).toBeGreaterThan(0)
      expect(skill.host.length, `${name}.host`).toBeGreaterThan(0)
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
    expect(names).toEqual([
      'add-chart',
      'add-formula',
      'set-cell-value',
      'sort-range',
      'xl-add-table',
      'xl-find-text',
    ])
  })

  it('returns PPT skills for the powerpoint host', () => {
    const skills = listSkillsForHost('powerpoint')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual([
      'add-slide',
      'duplicate-slide',
      'ppt-add-bullets',
      'ppt-add-image',
      'ppt-add-table',
      'ppt-add-text',
    ])
  })

  it('returns Word skills for the word host', () => {
    const skills = listSkillsForHost('word')
    const names = skills.map((s) => s.name).sort()
    expect(names).toEqual([
      'add-image',
      'add-page-break',
      'replace-text',
      'set-font',
      'word-add-table',
      'word-add-text',
      'word-find-text',
    ])
  })

  it('returns nothing for unsupported', () => {
    expect(listSkillsForHost('unsupported')).toEqual([])
  })
})
