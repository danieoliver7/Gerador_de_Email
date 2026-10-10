import type { EmailSettings } from './email'
import { DEFAULT_HEADER_SETTINGS, restoreHeaderSettings } from './header'

export const SETTINGS_STORAGE_KEY = 'atelie-email.settings.v1'

export const FONT_OPTIONS = [
  { label: 'Arial', value: 'Arial, Helvetica, sans-serif' },
  { label: 'Georgia', value: 'Georgia, serif' },
  { label: 'Verdana', value: 'Verdana, Geneva, sans-serif' },
  { label: 'Tahoma', value: 'Tahoma, Geneva, sans-serif' },
] as const

export const DEFAULT_SETTINGS: EmailSettings = {
  width: 600,
  fontFamily: FONT_OPTIONS[0].value,
  backgroundColor: '#f1f4f8',
  surfaceColor: '#ffffff',
  textColor: '#273344',
  accentColor: '#635bdb',
  template: 'classic',
  header: { ...DEFAULT_HEADER_SETTINGS },
}

const isHexColor = (value: unknown): value is string =>
  typeof value === 'string' && /^#[0-9a-fA-F]{6}$/.test(value)

export function loadSettings(): EmailSettings {
  try {
    const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY)
    if (!raw) return DEFAULT_SETTINGS
    const value: unknown = JSON.parse(raw)
    if (!value || typeof value !== 'object') return DEFAULT_SETTINGS
    const saved = value as Record<string, unknown>
    return {
      width: typeof saved.width === 'number' && Number.isInteger(saved.width)
        ? Math.max(480, Math.min(760, saved.width))
        : DEFAULT_SETTINGS.width,
      fontFamily: FONT_OPTIONS.find((font) => font.value === saved.fontFamily)?.value
        ?? DEFAULT_SETTINGS.fontFamily,
      backgroundColor: isHexColor(saved.backgroundColor) ? saved.backgroundColor : DEFAULT_SETTINGS.backgroundColor,
      surfaceColor: isHexColor(saved.surfaceColor) ? saved.surfaceColor : DEFAULT_SETTINGS.surfaceColor,
      textColor: isHexColor(saved.textColor) ? saved.textColor : DEFAULT_SETTINGS.textColor,
      accentColor: isHexColor(saved.accentColor) ? saved.accentColor : DEFAULT_SETTINGS.accentColor,
      template: saved.template === 'compact' ? 'compact' : 'classic',
      header: restoreHeaderSettings(saved.header),
    }
  } catch {
    return DEFAULT_SETTINGS
  }
}

export function saveSettings(settings: EmailSettings): boolean {
  try {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings))
    return true
  } catch {
    return false
  }
}
