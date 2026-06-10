/**
 * Tests for the tools layer (skill → AI SDK 5 Tool conversion).
 */
import { describe, expect, it } from 'vitest'
import { toolsForHost } from './tools'

describe('toolsForHost', () => {
  it('returns an empty object for hosts with no skills', () => {
    expect(toolsForHost('unsupported')).toEqual({})
  })

  it('returns 11 Excel skills', () => {
    const tools = toolsForHost('excel')
    const names = Object.keys(tools).sort()
    expect(names).toEqual([
      'add-chart',
      'add-formula',
      'format-chart',
      'set-cell-value',
      'set-chart-axis-title',
      'set-chart-data-labels',
      'set-chart-legend',
      'set-chart-title',
      'sort-range',
      'xl-add-table',
      'xl-find-text',
    ])
  })

  it('returns 6 PowerPoint skills', () => {
    const tools = toolsForHost('powerpoint')
    const names = Object.keys(tools).sort()
    expect(names).toEqual([
      'add-slide',
      'duplicate-slide',
      'ppt-add-bullets',
      'ppt-add-image',
      'ppt-add-table',
      'ppt-add-text',
    ])
  })

  it('returns 7 Word skills', () => {
    const tools = toolsForHost('word')
    const names = Object.keys(tools).sort()
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

  it('forwards executor errors from skill.execute as tool rejections', async () => {
    const tools = toolsForHost('excel')
    const tool = tools['add-chart']!
    const exec = (tool as unknown as { execute: (input: unknown) => Promise<unknown> }).execute
    await expect(
      exec({ sheetName: 'Sheet1', chartType: 'column-clustered', dataRange: 'A1:B5' }),
    ).rejects.toThrow(/Excel namespace not available/)
  })

  it('rejects malformed args at the Zod parse boundary', async () => {
    const tools = toolsForHost('excel')
    const tool = tools['add-chart']!
    const exec = (tool as unknown as { execute: (input: unknown) => Promise<unknown> }).execute
    await expect(
      exec({ sheetName: 'Sheet1', chartType: 'octopus' }),
    ).rejects.toThrow()
  })
})
