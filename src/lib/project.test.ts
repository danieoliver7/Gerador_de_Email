import { afterEach, describe, expect, it, vi } from 'vitest'
import type { ExtractedDocument } from './document'
import { renderEmail, type EmailSettings } from './email'
import type { ComponentOverrides } from './custom-components'
import { DEFAULT_HEADER_SETTINGS } from './header'
import {
  PROJECT_STORAGE_KEY,
  clearProjectStorage,
  createProjectData,
  loadProjectFromStorage,
  parseProject,
  saveProjectToStorage,
  serializeProject,
} from './project'

const settings: EmailSettings = {
  width: 600,
  fontFamily: 'Arial, Helvetica, sans-serif',
  backgroundColor: '#f1f4f8',
  surfaceColor: '#ffffff',
  textColor: '#273344',
  accentColor: '#635bdb',
  template: 'classic',
  header: { ...DEFAULT_HEADER_SETTINGS },
}

const document: ExtractedDocument = {
  sourceName: 'campanha.docx',
  warnings: ['Aviso original.'],
  blocks: [
    { type: 'paragraph', runs: [{ text: 'Antes ' }, { text: 'forte', bold: true }] },
    { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
    { type: 'paragraph', runs: [{ text: 'Entre ', italic: true }, { text: 'site', href: 'https://example.com' }] },
    { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-2' },
    { type: 'paragraph', runs: [{ text: 'Depois <script>é texto</script>' }] },
  ],
}

const overrides: Record<string, ComponentOverrides> = {
  'acessar-1': { text: 'Acessar agora', url: 'https://example.com/um?a=1&b=2' },
  'acessar-2': { text: 'Ver detalhes', url: 'https://example.com/dois' },
}

afterEach(() => {
  vi.restoreAllMocks()
  window.localStorage.clear()
})

describe('project serialization', () => {
  it('round-trips several component instances and preserves copy, order and settings', () => {
    const project = createProjectData(document, overrides, settings)
    const loaded = parseProject(serializeProject(project))

    expect(loaded.version).toBe(1)
    expect(loaded.document).toEqual(document)
    expect(loaded.document.blocks.map((block) => block.type)).toEqual([
      'paragraph', 'component', 'paragraph', 'component', 'paragraph',
    ])
    expect(loaded.componentOverrides).toEqual(overrides)
    expect(loaded.settings).toEqual(settings)
    expect(loaded.savedAt).toBe(project.savedAt)
    expect(serializeProject(project)).not.toContain('<table')
  })

  it('saves customized header settings and restores them unchanged', () => {
    const changed: EmailSettings = {
      ...settings,
      header: {
        enabled: false,
        backgroundColor: '#123456',
        height: 175,
        marginTop: 12,
        marginBottom: 24,
      },
    }
    const project = createProjectData(document, overrides, changed)
    expect(parseProject(serializeProject(project)).settings.header).toEqual(changed.header)
    expect(saveProjectToStorage(project)).toBe(true)
    expect(loadProjectFromStorage()?.settings.header).toEqual(changed.header)
  })

  it('migrates projects saved before header settings existed', () => {
    const legacy = JSON.parse(serializeProject(createProjectData(document, overrides, settings))) as {
      settings: Record<string, unknown>
    }
    delete legacy.settings.header
    const loaded = parseProject(JSON.stringify(legacy))
    expect(loaded.settings.header).toEqual(DEFAULT_HEADER_SETTINGS)
    expect(JSON.parse(serializeProject(loaded)).settings.header).toEqual(DEFAULT_HEADER_SETTINGS)
  })

  it('exports the same buttons after saving and reopening a project', async () => {
    const original = createProjectData(document, overrides, settings)
    const reopened = parseProject(serializeProject(original))
    const before = await renderEmail(original.document, original.settings, original.componentOverrides)
    const after = await renderEmail(reopened.document, reopened.settings, reopened.componentOverrides)
    expect(before.errors).toEqual([])
    expect(after.errors).toEqual([])
    expect(after.html).toBe(before.html)
    expect(after.html).toContain('Acessar agora')
    expect(after.html).toContain('Ver detalhes')
  })

  it('round-trips Word button labels and links, including a placeholder URL and a legacy pending URL, through JSON and storage', () => {
    const wordDocument: ExtractedDocument = {
      sourceName: 'botoes.docx',
      warnings: [],
      blocks: [
        { type: 'paragraph', runs: [{ text: 'Antes' }] },
        {
          type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
          initial: { text: 'BAIXAR AGORA', url: 'https://example.com/planilha?x=1&y=2' },
        },
        {
          type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-2',
          initial: { text: 'SAIBA MAIS', url: '#' },
        },
        {
          type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-3',
          initial: { text: 'LEGADO', url: '' },
        },
      ],
    }
    const project = createProjectData(wordDocument, {}, settings)
    const reopened = parseProject(serializeProject(project))

    expect(reopened.document).toEqual(wordDocument)
    expect(reopened.document.blocks[1]).toEqual(wordDocument.blocks[1])
    expect(reopened.document.blocks[2]).toEqual(wordDocument.blocks[2])
    expect(reopened.document.blocks[3]).toEqual(wordDocument.blocks[3])
    expect(saveProjectToStorage(project)).toBe(true)
    expect(loadProjectFromStorage()?.document).toEqual(wordDocument)
  })

  it('preserves conditional boundaries, copy and button edits after saving and reopening', async () => {
    const conditionalDocument: ExtractedDocument = {
      sourceName: 'audiencias.docx', warnings: [], errors: [],
      blocks: [
        { type: 'paragraph', runs: [{ text: 'Fora' }] },
        { type: 'condition-start', runs: [], audience: 'CLIENTES_VIP' },
        { type: 'paragraph', runs: [{ text: 'Dentro ', italic: true }, { text: 'forte', bold: true }] },
        { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
        { type: 'condition-end', runs: [], audience: 'CLIENTES_VIP' },
      ],
    }
    const changed = { 'acessar-1': { text: 'Ver oferta', url: 'https://example.org/vip' } }
    const original = createProjectData(conditionalDocument, changed, settings)
    const reopened = parseProject(serializeProject(original))
    expect(reopened.document).toEqual(conditionalDocument)
    expect(reopened.componentOverrides).toEqual(changed)
    expect(saveProjectToStorage(original)).toBe(true)
    expect(loadProjectFromStorage()?.document).toEqual(conditionalDocument)
    const before = await renderEmail(original.document, original.settings, original.componentOverrides)
    const after = await renderEmail(reopened.document, reopened.settings, reopened.componentOverrides)
    expect(after.errors).toEqual([])
    expect(after.html).toBe(before.html)
    expect(after.html).toContain('{%#if profile.segmentMembership.get("ups").get("CLIENTES_VIP").status = "existing"%}')
    expect(after.html).toContain('Ver oferta')
  })

  it('rejects malformed projects and unsafe document links without creating HTML', () => {
    const original = JSON.parse(serializeProject(createProjectData(document, overrides, settings))) as Record<string, unknown>
    const altered = (change: (copy: Record<string, unknown>) => void) => {
      const copy = structuredClone(original)
      change(copy)
      return JSON.stringify(copy)
    }

    expect(() => parseProject('not json')).toThrow('Projeto inválido')
    expect(() => parseProject(altered((copy) => { copy.version = 2 }))).toThrow('Projeto inválido')
    expect(() => parseProject(altered((copy) => {
      const blocks = (copy.document as ExtractedDocument).blocks
      blocks[2].runs[1].href = 'data:text/html,<script>alert(1)</script>'
    }))).toThrow('Projeto inválido')
    expect(() => parseProject(altered((copy) => {
      (copy.document as ExtractedDocument).blocks.push({
        type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1',
      })
    }))).toThrow('Projeto inválido')
    expect(() => parseProject(altered((copy) => { copy.html = '<script>alert(1)</script>' })))
      .toThrow('Projeto inválido')
    for (const header of [
      null,
      { ...DEFAULT_HEADER_SETTINGS, backgroundColor: 'url(javascript:alert(1))' },
      { ...DEFAULT_HEADER_SETTINGS, height: -1 },
      { ...DEFAULT_HEADER_SETTINGS, marginBottom: 9999 },
      { ...DEFAULT_HEADER_SETTINGS, enabled: 'yes' },
      { ...DEFAULT_HEADER_SETTINGS, unexpected: true },
    ]) {
      expect(() => parseProject(altered((copy) => {
        (copy.settings as Record<string, unknown>).header = header
      }))).toThrow('Projeto inválido')
    }
    expect(() => createProjectData({
      ...document,
      blocks: [{
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'ABRIR', url: 'javascript:alert(1)' },
      }],
    }, {}, settings)).toThrow('Projeto inválido')
    expect(() => createProjectData({
      ...document,
      blocks: [{ type: 'condition-start', runs: [], audience: 'VIP"%}' }],
    }, {}, settings)).toThrow('Projeto inválido')
  })

  it('keeps incomplete URL drafts inert for later editing', () => {
    const project = createProjectData(document, {
      ...overrides,
      'acessar-1': { text: 'Rascunho', url: 'https://' },
      'acessar-2': { text: 'Pendente', url: 'javascript:alert(1)' },
    }, settings)
    const serialized = serializeProject(project)
    expect(parseProject(serialized).componentOverrides['acessar-1'].url).toBe('https://')
    expect(parseProject(serialized).componentOverrides['acessar-2'].url).toBe('javascript:alert(1)')
    expect(serialized).not.toContain('href=')
    expect(() => createProjectData(document, {
      ...overrides,
      'acessar-1': { text: 'Inválido', url: 'https://example.com\n<script>' },
    }, settings)).toThrow('Projeto inválido')
  })
})

describe('local project storage', () => {
  it('saves, reloads and clears a project with its component edits', () => {
    const project = createProjectData(document, overrides, settings)
    expect(saveProjectToStorage(project)).toBe(true)
    expect(window.localStorage.getItem(PROJECT_STORAGE_KEY)).toBe(serializeProject(project))
    expect(loadProjectFromStorage()).toEqual(project)
    expect(clearProjectStorage()).toBe(true)
    expect(loadProjectFromStorage()).toBeNull()
  })

  it('fails safely on invalid stored data or unavailable storage', () => {
    window.localStorage.setItem(PROJECT_STORAGE_KEY, '{invalid')
    expect(loadProjectFromStorage()).toBeNull()

    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => { throw new Error('quota') })
    expect(saveProjectToStorage(createProjectData(document, overrides, settings))).toBe(false)
  })
})
