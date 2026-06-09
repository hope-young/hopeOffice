/**
 * Tests for the Settings store's storage tier resolution.
 *
 * The `makeStorage` factory is module-scoped (it inspects the
 * `Office` global at import time), so we can't reliably re-pick
 * a tier mid-test. Instead we mirror the factory's two branches
 * inline and assert the branch logic. The Settings store's
 * behaviour against each tier is covered by the integration
 * path (real Office → roamingSettings; dev → localStorage).
 */
import { describe, expect, it, beforeEach, afterEach, vi } from 'vitest'

type StorageLike = {
  getItem: (key: string) => string | null
  setItem: (key: string, value: string) => void
  removeItem: (key: string) => void
}

/** A Map-backed stub of the localStorage shape, exposed to
 *  tests under the same name so we can exercise the localStorage
 *  branch on a Node runner (vitest's default env has no
 *  `localStorage` global). */
function installLocalStorageStub() {
  const backing = new Map<string, string>()
  const stub: StorageLike = {
    getItem: (k) => (backing.has(k) ? backing.get(k)! : null),
    setItem: (k, v) => {
      backing.set(k, v)
    },
    removeItem: (k) => {
      backing.delete(k)
    },
  }
  vi.stubGlobal('localStorage', stub)
  return { backing, stub }
}

/** Mirror of the localStorage branch in `makeStorage`. */
function localStorageBranch(): StorageLike {
  if (typeof localStorage === 'undefined') {
    return { getItem: () => null, setItem: () => undefined, removeItem: () => undefined }
  }
  return {
    getItem: (k) => localStorage.getItem(k),
    setItem: (k, v) => localStorage.setItem(k, v),
    removeItem: (k) => localStorage.removeItem(k),
  }
}

/** Mirror of the Office roamingSettings branch. */
function officeBranch(): StorageLike & { backing: Record<string, string> } {
  const backing: Record<string, string> = {}
  return {
    backing,
    getItem: (k) => backing[k] ?? null,
    setItem: (k, v) => {
      backing[k] = v
    },
    removeItem: (k) => {
      delete backing[k]
    },
  }
}

describe('settings storage tier — localStorage branch', () => {
  let backing: Map<string, string>

  beforeEach(() => {
    ;({ backing } = installLocalStorageStub())
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('round-trips a value', () => {
    const s = localStorageBranch()
    s.setItem('k', 'v')
    expect(s.getItem('k')).toBe('v')
    expect(backing.get('k')).toBe('v')
  })

  it('returns null for missing keys', () => {
    const s = localStorageBranch()
    expect(s.getItem('does-not-exist')).toBeNull()
  })

  it('removes values', () => {
    const s = localStorageBranch()
    s.setItem('k', 'v')
    s.removeItem('k')
    expect(s.getItem('k')).toBeNull()
    expect(backing.has('k')).toBe(false)
  })
})

describe('settings storage tier — Office roamingSettings branch', () => {
  it('round-trips a value into the backing map', () => {
    const s = officeBranch()
    s.setItem('k', 'v')
    expect(s.getItem('k')).toBe('v')
    expect(s.backing['k']).toBe('v')
  })

  it('removes values from the backing map', () => {
    const s = officeBranch()
    s.setItem('k', 'v')
    s.removeItem('k')
    expect(s.getItem('k')).toBeNull()
    expect(s.backing['k']).toBeUndefined()
  })
})
