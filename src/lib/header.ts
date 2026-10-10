export type HeaderSettings = {
  enabled: boolean
  backgroundColor: string
  height: number
  marginTop: number
  marginBottom: number
}

export const HEADER_HEIGHT_MIN = 20
export const HEADER_HEIGHT_MAX = 600
export const HEADER_MARGIN_MIN = 0
export const HEADER_MARGIN_MAX = 200

export const DEFAULT_HEADER_SETTINGS: HeaderSettings = {
  enabled: true,
  backgroundColor: '#000000',
  height: 100,
  marginTop: 0,
  marginBottom: 0,
}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  value !== null && typeof value === 'object' && !Array.isArray(value)

const isHexColor = (value: unknown): value is string =>
  typeof value === 'string' && /^#[0-9a-fA-F]{6}$/.test(value)

const isIntegerInRange = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isInteger(value) && value >= min && value <= max

export function isValidHeaderSettings(value: unknown): value is HeaderSettings {
  if (!isRecord(value)) return false
  const keys = Object.keys(value)
  if (keys.length !== 5 || ![
    'enabled', 'backgroundColor', 'height', 'marginTop', 'marginBottom',
  ].every((key) => Object.hasOwn(value, key))) return false
  return typeof value.enabled === 'boolean' &&
    isHexColor(value.backgroundColor) &&
    isIntegerInRange(value.height, HEADER_HEIGHT_MIN, HEADER_HEIGHT_MAX) &&
    isIntegerInRange(value.marginTop, HEADER_MARGIN_MIN, HEADER_MARGIN_MAX) &&
    isIntegerInRange(value.marginBottom, HEADER_MARGIN_MIN, HEADER_MARGIN_MAX)
}

/** Tolerant loading for browser settings saved before the header existed. */
export function restoreHeaderSettings(value: unknown): HeaderSettings {
  if (!isRecord(value)) return { ...DEFAULT_HEADER_SETTINGS }
  return {
    enabled: typeof value.enabled === 'boolean' ? value.enabled : DEFAULT_HEADER_SETTINGS.enabled,
    backgroundColor: isHexColor(value.backgroundColor)
      ? value.backgroundColor : DEFAULT_HEADER_SETTINGS.backgroundColor,
    height: isIntegerInRange(value.height, HEADER_HEIGHT_MIN, HEADER_HEIGHT_MAX)
      ? value.height : DEFAULT_HEADER_SETTINGS.height,
    marginTop: isIntegerInRange(value.marginTop, HEADER_MARGIN_MIN, HEADER_MARGIN_MAX)
      ? value.marginTop : DEFAULT_HEADER_SETTINGS.marginTop,
    marginBottom: isIntegerInRange(value.marginBottom, HEADER_MARGIN_MIN, HEADER_MARGIN_MAX)
      ? value.marginBottom : DEFAULT_HEADER_SETTINGS.marginBottom,
  }
}
