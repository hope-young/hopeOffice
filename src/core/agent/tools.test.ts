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
    expect(toolsForHost('word')).toEqual({})
    expect(toolsForHost('unsupported')).toEqual({})
  })

  it('returns add-chart for excel', () => {
    const tools = toolsForHost('excel')
    expect(Object.keys(tools).sort()).toEqual(['add-chart'])
    expect(tools['add-chart']?.description).toContain('chart')
  })

  it('returns add-text + add-slide for powerpoint', () => {
    const tools = toolsForHost('powerpoint')
    expect(Object.keys(tools).sort()).toEqual(['add-slide', 'add-text'])
    expect(tools['add-text']?.description.toLowerCase()).toContain('text')
    expect(tools['add-slide']?.description.toLowerCase()).toContain('slide')
  })

  it('forwards executor errors from skill.execute as tool rejections', async () => {
    // Use excel's add-chart but in a context where the underlying
    // namespace is missing. The skill itself throws synchronously
    // before office.js is touched.
    const tools = toolsForHost('excel')
    const tool = tools['add-chart']
    expect(tool).toBeDefined()
    if (!tool) return

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
    const tool = tools['add-chart']
    if (!tool) throw new Error('add-chart tool missing')
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
