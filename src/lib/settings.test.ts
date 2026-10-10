import { afterEach, describe, expect, it } from 'vitest'
import { DEFAULT_HEADER_SETTINGS } from './header'
import { DEFAULT_SETTINGS, SETTINGS_STORAGE_KEY, loadSettings, saveSettings } from './settings'

afterEach(() => window.localStorage.clear())

describe('email settings storage', () => {
  it('enables the black, 100 px header by default', () => {
    expect(loadSettings().header).toEqual(DEFAULT_HEADER_SETTINGS)
  })

  it('persists the header together with existing email settings', () => {
    const changed = {
      ...DEFAULT_SETTINGS,
      width: 640,
      header: {
        enabled: false,
        backgroundColor: '#123456',
        height: 140,
        marginTop: 8,
        marginBottom: 16,
      },
    }
    expect(saveSettings(changed)).toBe(true)
    expect(loadSettings()).toEqual(changed)
  })

  it('adds defaults to browser settings saved before the header existed', () => {
    const legacy = { ...DEFAULT_SETTINGS } as Record<string, unknown>
    delete legacy.header
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(legacy))
    expect(loadSettings().header).toEqual(DEFAULT_HEADER_SETTINGS)
    expect(loadSettings().width).toBe(DEFAULT_SETTINGS.width)
  })

  it('restores valid fields while rejecting malformed stored header fields', () => {
    const stored = {
      ...DEFAULT_SETTINGS,
      header: {
        enabled: false,
        backgroundColor: 'not a color',
        height: -20,
        marginTop: 24,
        marginBottom: 1000,
      },
    }
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(stored))
    expect(loadSettings().header).toEqual({
      enabled: false,
      backgroundColor: '#000000',
      height: 100,
      marginTop: 24,
      marginBottom: 0,
    })
  })
})
