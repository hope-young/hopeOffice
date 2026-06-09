/**
 * Tests for the tools layer (skill → AI SDK 5 Tool conversion).
 * The execute function is exercised by stubbing the skill body —
 * we don't have office.js in jsdom, so any code path that calls
 * into PowerPoint.run() would throw, which is the same behavior
 * the dev-browser path surfaces.
 */
import { describe, expect, it, vi } from 'vitest'
import { toolsForHost } from './tools'

describe('toolsForHost', () => {
  it('returns an empty object for hosts with no skills', () => {
    expect(toolsForHost('unsupported')).toEqual({})
  })

  it('returns add-chart for excel', () => {
    const tools = toolsForHost('excel')
    expect(Object.keys(tools).sort()).toEqual(['add-chart'])
    expect(tools['add-chart']?.description).toContain('chart')
  })

  it('returns ppt-add-text + add-slide for powerpoint', () => {
    const tools = toolsForHost('powerpoint')
    expect(Object.keys(tools).sort()).toEqual(['add-slide', 'ppt-add-text'])
    const pptAddText = tools['ppt-add-text']!
    expect((pptAddText.description ?? '').toLowerCase()).toContain('text')
    const addSlide = tools['add-slide']!
    expect((addSlide.description ?? '').toLowerCase()).toContain('slide')
  })

  it('returns add-table + word-add-text for word', () => {
    const tools = toolsForHost('word')
    expect(Object.keys(tools).sort()).toEqual(['add-table', 'word-add-text'])
    const addTable = tools['add-table']!
    expect((addTable.description ?? '').toLowerCase()).toContain('table')
    const wordAddText = tools['word-add-text']!
    expect((wordAddText.description ?? '').toLowerCase()).toContain('text')
  })

  it('forwards executor errors from skill.execute as tool rejections', async () => {
    // Use excel's add-chart but in a context where the underlying
    // namespace is missing. The skill itself throws synchronously
    // before office.js is touched.
    const tools = toolsForHost('excel')
    const tool = tools['add-chart']!

    // The execute function lives on the AI SDK Tool object.
    // Type assertion: we know our `execute` returns Promise<unknown>
    // because we wrote it that way in tools.ts.
    const exec = (tool as unknown as { execute: (input: unknown) => Promise<unknown> }).execute
    await expect(
      exec({ sheetName: 'Sheet1', chartType: 'column-clustered', dataRange: 'A1:B5' }),
    ).rejects.toThrow(/Excel namespace not available/)
  })

  it('rejects malformed args at the Zod parse boundary', async () => {
    const tools = toolsForHost('excel')
    const tool = tools['add-chart']!
    const exec = (tool as unknown as { execute: (input: unknown) => Promise<unknown> }).execute
    // Zod parse failure should bubble up as a thrown error.
    await expect(
      // chartType is wrong (not in enum) + dataRange missing.
      exec({ sheetName: 'Sheet1', chartType: 'octopus' }),
    ).rejects.toThrow()
  })
})

// Reference vi for future use; silences the unused import warning
// when no future test actually needs a mock.
void vi
