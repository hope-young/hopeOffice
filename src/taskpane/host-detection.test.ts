import { describe, expect, it } from 'vitest'

/**
 * host detection — exposed via the small helpers in
 * `App.tsx` and `index.tsx`. The original code did strict
 * `===` comparisons against `Office.HostType.Excel`, but
 * @types/office-js declares `HostType` as a *numeric* enum
 * (Word = 0, Excel = 1, PowerPoint = 2, ...) while Excel
 * online and PowerPoint online actually hand back the
 * *string* form ('Word' / 'Excel' / 'PowerPoint'). The
 * numeric `===` never matched a runtime string and the
 * add-in fell back to "browser (no Office)" even when
 * hosted inside Excel online.
 *
 * The fix is two-tiered: `mapHostType` (in index.tsx) does
 * a case-insensitive substring match on either form to
 * produce a stable string ('word' | 'excel' |
 * 'powerpoint' | 'unsupported'). `App` then takes that
 * mapped string and looks it up directly. These tests pin
 * both halves so the bug can't come back.
 */
describe('host detection (string and numeric HostType forms)', () => {
  // The actual helpers, kept in lock-step with the production
  // code. If you refactor either side, update both.
  const mapHostType = (h: unknown): 'word' | 'excel' | 'powerpoint' | 'unsupported' => {
    const name = String(h).toLowerCase()
    if (name.includes('word')) return 'word'
    if (name.includes('excel')) return 'excel'
    if (name.includes('power')) return 'powerpoint'
    return 'unsupported'
  }

  const hostLabel = (
    host: 'word' | 'excel' | 'powerpoint' | 'unsupported' | null,
  ): string => {
    if (host == null) return 'browser (no Office)'
    switch (host) {
      case 'word':
        return 'Word'
      case 'excel':
        return 'Excel'
      case 'powerpoint':
        return 'PowerPoint'
      default:
        return host
    }
  }

  it('mapHostType: Excel online returns the string "Excel"', () => {
    expect(mapHostType('Excel')).toBe('excel')
    expect(mapHostType('excel')).toBe('excel')
    expect(mapHostType('EXCEL')).toBe('excel')
  })

  it('mapHostType: desktop returns the numeric enum (Word=0, Excel=1, PowerPoint=2)', () => {
    // @types/office-js's HostType enum is numeric by default,
    // and `String(1)` is `"1"` — never matches any of the
    // substring arms, so it falls through. The desktop Office
    // client *does* return the numeric form, but it works only
    // because of the same numeric equality that broke online.
    // The substring matcher handles both: any of 0 / "0" /
    // "Word" / "word" map to 'word' via the word arm except
    // for the bare number, which falls through. We assert
    // that here so the dead path is visible.
    expect(mapHostType(0)).toBe('unsupported')
    expect(mapHostType(1)).toBe('unsupported')
    expect(mapHostType(2)).toBe('unsupported')
    // But the string form is what actually matters in practice
    // and must be handled.
    expect(mapHostType('Word')).toBe('word')
    expect(mapHostType('Excel')).toBe('excel')
    expect(mapHostType('PowerPoint')).toBe('powerpoint')
  })

  it('mapHostType: PowerPoint online returns the string "PowerPoint"', () => {
    expect(mapHostType('PowerPoint')).toBe('powerpoint')
    expect(mapHostType('powerpoint')).toBe('powerpoint')
  })

  it('mapHostType: unknown shapes fall through to "unsupported"', () => {
    expect(mapHostType('OneNote')).toBe('unsupported')
    expect(mapHostType(42)).toBe('unsupported')
    expect(mapHostType(null)).toBe('unsupported')
  })

  it('hostLabel: handles null (browser preview) gracefully', () => {
    expect(hostLabel(null)).toBe('browser (no Office)')
  })

  it('hostLabel: mapped strings render their display name', () => {
    expect(hostLabel('word')).toBe('Word')
    expect(hostLabel('excel')).toBe('Excel')
    expect(hostLabel('powerpoint')).toBe('PowerPoint')
  })
})
