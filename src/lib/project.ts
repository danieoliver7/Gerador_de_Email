import type { ExtractedDocument } from './document'
import type { EmailSettings } from './email'
import type { ComponentOverrides } from './custom-components'
import { validAudience } from './conditions'
import { DEFAULT_HEADER_SETTINGS, isValidHeaderSettings } from './header'

export type ProjectData = {
  version: 1
  document: ExtractedDocument
  componentOverrides: Record<string, ComponentOverrides>
  settings: EmailSettings
  savedAt: string
}

export const PROJECT_STORAGE_KEY = 'atelie-email.project.v1'

const MAX_PROJECT_LENGTH = 10_000_000
const MAX_BLOCKS = 20_000
const MAX_RUNS = 20_000
const MAX_STRING_LENGTH = 2_000_000
const MAX_BUTTON_TEXT_LENGTH = 1_000
const MAX_URL_LENGTH = 4_096

function invalidProject(): never {
  throw new Error('Projeto inválido ou incompatível.')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false
  const prototype: unknown = Object.getPrototypeOf(value)
  return prototype === Object.prototype || prototype === null
}

function hasKeys(value: Record<string, unknown>, required: string[], optional: string[] = []): boolean {
  const allowed = new Set([...required, ...optional])
  return required.every((key) => Object.hasOwn(value, key)) &&
    Object.keys(value).every((key) => allowed.has(key))
}

function isText(value: unknown, maximum = MAX_STRING_LENGTH): value is string {
  return typeof value === 'string' && value.length <= maximum
}

function isSafeUrl(value: unknown): value is string {
  if (!isText(value, MAX_URL_LENGTH)) return false
  if (value === '' || /\s|[\u0000-\u001f\u007f]/.test(value)) return false
  if (/^#[A-Za-z0-9._-]+$/.test(value)) return true
  if (/^https?:\/\//i.test(value)) {
    try {
      const parsed = new URL(value)
      return Boolean(parsed.hostname) && (parsed.protocol === 'http:' || parsed.protocol === 'https:')
    } catch {
      return false
    }
  }
  return /^(mailto|tel):[^\s]+$/i.test(value)
}

// An override is draft input. Preview/export validate it before creating a link;
// keeping incomplete input here lets the editor restore an unfinished project.
function isDraftUrl(value: unknown): value is string {
  return isText(value, MAX_URL_LENGTH) && !/[\u0000-\u001f\u007f-\u009f]/.test(value)
}

function validateComponentInitial(value: unknown): ComponentOverrides {
  if (!isRecord(value) || !hasKeys(value, ['text', 'url']) ||
    !isText(value.text, MAX_BUTTON_TEXT_LENGTH) ||
    /[\u0000-\u001f\u007f-\u009f]/.test(value.text) ||
    !(value.url === '' || value.url === '#' || (isSafeUrl(value.url) && /^https?:\/\//i.test(value.url)))) {
    return invalidProject()
  }
  return { text: value.text, url: value.url }
}

function validateRun(value: unknown): ExtractedDocument['blocks'][number]['runs'][number] {
  if (!isRecord(value) || !hasKeys(value, ['text'], ['bold', 'italic', 'href']) || !isText(value.text)) {
    return invalidProject()
  }
  if (value.bold !== undefined && typeof value.bold !== 'boolean') return invalidProject()
  if (value.italic !== undefined && typeof value.italic !== 'boolean') return invalidProject()
  if (value.href !== undefined && !isSafeUrl(value.href)) return invalidProject()
  return {
    text: value.text,
    ...(value.bold === undefined ? {} : { bold: value.bold }),
    ...(value.italic === undefined ? {} : { italic: value.italic }),
    ...(value.href === undefined ? {} : { href: value.href }),
  }
}

function validateBlock(value: unknown): ExtractedDocument['blocks'][number] {
  if (!isRecord(value) || typeof value.type !== 'string') return invalidProject()

  if (value.type === 'condition-start' || value.type === 'condition-end') {
    if (!hasKeys(value, ['type', 'runs', 'audience']) ||
      !Array.isArray(value.runs) || value.runs.length !== 0 ||
      typeof value.audience !== 'string' || !validAudience(value.audience)) return invalidProject()
    return { type: value.type, runs: [], audience: value.audience }
  }

  if (value.type === 'component') {
    if (!hasKeys(value, ['type', 'runs', 'componentId', 'instanceId'], ['initial']) ||
      !Array.isArray(value.runs) || value.runs.length !== 0 ||
      !isText(value.componentId, 80) || !/^[a-z][a-z0-9_-]*$/.test(value.componentId) ||
      !isText(value.instanceId, 100) ||
      !value.instanceId.startsWith(`${value.componentId}-`) ||
      !/^[a-z][a-z0-9_-]*-[1-9][0-9]*$/.test(value.instanceId)) {
      return invalidProject()
    }
    return {
      type: 'component',
      runs: [],
      componentId: value.componentId,
      instanceId: value.instanceId,
      ...(value.initial === undefined ? {} : { initial: validateComponentInitial(value.initial) }),
    } as ExtractedDocument['blocks'][number]
  }

  if (value.type !== 'paragraph' && value.type !== 'heading' && value.type !== 'list-item') return invalidProject()
  const required = value.type === 'heading'
    ? ['type', 'runs', 'level']
    : value.type === 'list-item'
      ? ['type', 'runs', 'ordered', 'level']
      : ['type', 'runs']
  if (!hasKeys(value, required) || !Array.isArray(value.runs) || value.runs.length > MAX_RUNS) {
    return invalidProject()
  }
  if (value.type === 'heading' && (!Number.isInteger(value.level) || Number(value.level) < 1 || Number(value.level) > 6)) {
    return invalidProject()
  }
  if (value.type === 'list-item' && (
    typeof value.ordered !== 'boolean' ||
    !Number.isInteger(value.level) || Number(value.level) < 1 || Number(value.level) > 100
  )) return invalidProject()

  const runs = value.runs.map(validateRun)
  if (value.type === 'heading') return { type: 'heading', runs, level: Number(value.level) }
  if (value.type === 'list-item') {
    return { type: 'list-item', runs, ordered: Boolean(value.ordered), level: Number(value.level) }
  }
  return { type: 'paragraph', runs }
}

function validateDocument(value: unknown): ExtractedDocument {
  if (!isRecord(value) || !hasKeys(value, ['blocks', 'warnings', 'sourceName'], ['errors']) ||
    !Array.isArray(value.blocks) || value.blocks.length > MAX_BLOCKS ||
    !Array.isArray(value.warnings) || value.warnings.length > 1_000 ||
    !value.warnings.every((warning) => isText(warning, 5_000)) ||
    (value.errors !== undefined && (!Array.isArray(value.errors) || value.errors.length > 1_000 ||
      !value.errors.every((error) => isText(error, 5_000)))) ||
    !isText(value.sourceName, 1_024)) return invalidProject()

  const blocks = value.blocks.map(validateBlock)
  const instanceIds = new Set<string>()
  for (const block of blocks) {
    if (block.type !== 'component') continue
    if (instanceIds.has(block.instanceId)) return invalidProject()
    instanceIds.add(block.instanceId)
  }
  return {
    blocks, warnings: [...value.warnings], sourceName: value.sourceName,
    ...(value.errors === undefined ? {} : { errors: [...value.errors] }),
  }
}

function validateSettings(value: unknown): EmailSettings {
  if (!isRecord(value) || !hasKeys(value, [
    'width', 'fontFamily', 'backgroundColor', 'surfaceColor', 'textColor', 'accentColor', 'template',
  ], ['header'])) return invalidProject()
  if (!Number.isInteger(value.width) || Number(value.width) < 320 || Number(value.width) > 800 ||
    !isText(value.fontFamily, 120) || !value.fontFamily.trim() ||
    !/^[\w\s,'"-]+$/.test(value.fontFamily) ||
    ![value.backgroundColor, value.surfaceColor, value.textColor, value.accentColor]
      .every((color) => typeof color === 'string' && /^#[0-9a-fA-F]{6}$/.test(color)) ||
    (value.template !== 'classic' && value.template !== 'compact')) return invalidProject()

  // Projects created before headers share the v1 format and gain its defaults.
  // A present header must be valid so imported JSON cannot bypass editor limits.
  const hasHeader = Object.hasOwn(value, 'header')
  if (hasHeader && !isValidHeaderSettings(value.header)) return invalidProject()

  return {
    width: Number(value.width),
    fontFamily: value.fontFamily,
    backgroundColor: value.backgroundColor as string,
    surfaceColor: value.surfaceColor as string,
    textColor: value.textColor as string,
    accentColor: value.accentColor as string,
    template: value.template,
    header: hasHeader ? value.header as EmailSettings['header'] : { ...DEFAULT_HEADER_SETTINGS },
  }
}

function validateOverrides(value: unknown, document: ExtractedDocument): Record<string, ComponentOverrides> {
  if (!isRecord(value)) return invalidProject()
  const instanceIds = new Set(
    document.blocks.filter((block) => block.type === 'component').map((block) => block.instanceId),
  )
  const overrides: Record<string, ComponentOverrides> = Object.create(null) as Record<string, ComponentOverrides>
  for (const [instanceId, entry] of Object.entries(value)) {
    if (!instanceIds.has(instanceId) || !isRecord(entry) || !hasKeys(entry, ['text', 'url']) ||
      !isText(entry.text, MAX_BUTTON_TEXT_LENGTH) || !isDraftUrl(entry.url)) return invalidProject()
    overrides[instanceId] = { text: entry.text, url: entry.url }
  }
  return overrides
}

function validateProject(value: unknown): ProjectData {
  if (!isRecord(value) || !hasKeys(value, [
    'version', 'document', 'componentOverrides', 'settings', 'savedAt',
  ]) || value.version !== 1 || !isText(value.savedAt, 30) ||
    Number.isNaN(Date.parse(value.savedAt)) || new Date(value.savedAt).toISOString() !== value.savedAt) {
    return invalidProject()
  }
  const document = validateDocument(value.document)
  return {
    version: 1,
    document,
    componentOverrides: validateOverrides(value.componentOverrides, document),
    settings: validateSettings(value.settings),
    savedAt: value.savedAt,
  }
}

export function createProjectData(
  document: ExtractedDocument,
  componentOverrides: Record<string, ComponentOverrides>,
  settings: EmailSettings,
): ProjectData {
  return validateProject({ version: 1, document, componentOverrides, settings, savedAt: new Date().toISOString() })
}

export function serializeProject(project: ProjectData): string {
  const serialized = JSON.stringify(validateProject(project))
  if (serialized.length > MAX_PROJECT_LENGTH) return invalidProject()
  return serialized
}

export function parseProject(raw: string): ProjectData {
  if (typeof raw !== 'string' || raw.length > MAX_PROJECT_LENGTH) return invalidProject()
  try {
    return validateProject(JSON.parse(raw) as unknown)
  } catch {
    return invalidProject()
  }
}

export function loadProjectFromStorage(): ProjectData | null {
  try {
    const raw = window.localStorage.getItem(PROJECT_STORAGE_KEY)
    return raw === null ? null : parseProject(raw)
  } catch {
    return null
  }
}

export function saveProjectToStorage(project: ProjectData): boolean {
  try {
    window.localStorage.setItem(PROJECT_STORAGE_KEY, serializeProject(project))
    return true
  } catch {
    return false
  }
}

export function clearProjectStorage(): boolean {
  try {
    window.localStorage.removeItem(PROJECT_STORAGE_KEY)
    return true
  } catch {
    return false
  }
}
