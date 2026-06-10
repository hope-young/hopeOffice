/**
 * Tests for the system prompt — focus on the bits that
 * change based on the live selection so we don't regress
 * the contract the LLM relies on.
 */
import { describe, expect, it } from 'vitest'
import { buildSystemPrompt } from './system-prompt'

describe('buildSystemPrompt', () => {
  it('includes the active-selection line so the LLM can route tool calls', () => {
    const prompt = buildSystemPrompt('excel', { kind: 'none' })
    expect(prompt).toContain('Active selection: none.')
  })

  it('renders a range selection with address + size', () => {
    const prompt = buildSystemPrompt('excel', {
      kind: 'range',
      address: 'A1:B2',
      rows: 2,
      cols: 2,
      preview: [
        [1, 2],
        [3, 4],
      ],
    })
    expect(prompt).toContain('Active selection: range A1:B2 (2 rows × 2 cols).')
  })

  it('flags a 1x1 selection as too small for add-chart', () => {
    const prompt = buildSystemPrompt('excel', {
      kind: 'range',
      address: 'A1',
      rows: 1,
      cols: 1,
      preview: [['x']],
    })
    expect(prompt).toContain('Active selection: single cell A1.')
    expect(prompt).toContain('needs ≥2 rows')
  })

  it('renders a chart selection with title', () => {
    const prompt = buildSystemPrompt('excel', {
      kind: 'chart',
      title: 'Q4 Revenue',
    })
    expect(prompt).toContain('Active selection: chart titled "Q4 Revenue".')
  })

  it('still works when no selection is provided (backwards compat)', () => {
    const prompt = buildSystemPrompt('excel')
    expect(prompt).toContain('Active selection: none.')
  })

  it('lists skills for the host', () => {
    const prompt = buildSystemPrompt('excel')
    // add-chart is the most basic Excel skill; if this is gone
    // something is very wrong.
    expect(prompt).toContain('add-chart')
  })
})
