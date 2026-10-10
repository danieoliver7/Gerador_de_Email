import { describe, expect, it } from 'vitest'
import type { ExtractedDocument } from './document'
import { buildMjml, renderEmail, type EmailSettings } from './email'
import { ACESSAR_HTML, renderCustomComponent } from './custom-components'
import { DEFAULT_HEADER_SETTINGS } from './header'
import { parseMammothHtml } from './document'

const settings: EmailSettings = {
  width: 600,
  fontFamily: 'Arial, Helvetica, sans-serif',
  backgroundColor: '#eef2f7',
  surfaceColor: '#ffffff',
  textColor: '#172033',
  accentColor: '#2054d2',
  template: 'classic',
  header: { ...DEFAULT_HEADER_SETTINGS },
}

function documentWith(...blocks: ExtractedDocument['blocks']): ExtractedDocument {
  return { blocks, sourceName: 'copy.docx', warnings: [] }
}

describe('buildMjml', () => {
  it('preserves paragraph order and inline formatting without treating Word text as HTML', () => {
    const document = documentWith(
      { type: 'heading', level: 1, runs: [{ text: 'Olá & boas-vindas' }] },
      {
        type: 'paragraph',
        runs: [
          { text: 'Texto ' },
          { text: 'forte', bold: true },
          { text: ' e ' },
          { text: 'ênfase', italic: true },
          { text: ' <script>alert(1)</script>' },
        ],
      },
    )

    const mjml = buildMjml(document, settings)
    expect(mjml.indexOf('Olá &amp; boas-vindas')).toBeLessThan(mjml.indexOf('Texto '))
    expect(mjml).toContain('<strong>forte</strong>')
    expect(mjml).toContain('<em>ênfase</em>')
    expect(mjml).toContain('&lt;script&gt;alert(1)&lt;/script&gt;')
    expect(mjml).not.toContain('<script>')
  })

  it('keeps safe Word links and does not create a link from unlinked copy', () => {
    const document = documentWith({
      type: 'paragraph',
      runs: [
        { text: 'Leia mais', href: 'https://example.com/campanha?a=1&b=2' },
        { text: ' Visite example.org' },
        { text: ' bloqueado', href: 'javascript:alert(1)' },
      ],
    })

    const mjml = buildMjml(document, settings)
    expect(mjml).toContain('href="https://example.com/campanha?a=1&amp;b=2"')
    expect(mjml.match(/<a\s/g)).toHaveLength(1)
    expect(mjml).not.toContain('href="javascript:')
    expect(mjml).toContain('Visite example.org')
  })

  it('uses the chosen width, colors and template without changing the copy', () => {
    const document = documentWith({ type: 'paragraph', runs: [{ text: 'Mesmo texto' }] })
    const classic = buildMjml(document, settings)
    const compact = buildMjml(document, { ...settings, template: 'compact' })

    expect(classic).toContain('width="600px"')
    expect(classic).toContain('background-color="#eef2f7"')
    expect(classic).toContain('background-color="#ffffff"')
    expect(classic).toContain('color="#172033"')
    expect(classic).toContain('font-family="Arial, Helvetica, sans-serif"')
    expect(classic).toContain('padding="24px 16px 0"')
    expect(compact).toContain('padding="0"')
    expect(compact).toContain('Mesmo texto')
    expect(classic).toContain('Mesmo texto')
  })

  it('places one automatic header after leading Word subject and fine line, including empty paragraphs', () => {
    const document = documentWith(
      { type: 'paragraph', runs: [{ text: 'Assunto: ' }, { text: 'Oferta', bold: true }] },
      { type: 'paragraph', runs: [{ text: '  ' }] },
      { type: 'paragraph', runs: [{ text: 'Linha fina: Confira os detalhes' }] },
      { type: 'paragraph', runs: [{ text: 'Corpo principal' }] },
    )
    const mjml = buildMjml(document, settings)
    const header = mjml.indexOf('bgcolor="#000000"')
    expect(header).toBeGreaterThan(mjml.indexOf('Assunto:'))
    expect(header).toBeGreaterThan(mjml.indexOf('Linha fina:'))
    expect(header).toBeLessThan(mjml.indexOf('Corpo principal'))
    expect(mjml.match(/bgcolor="#000000"/g)).toHaveLength(1)
    expect(mjml).toContain('<strong>Oferta</strong>')
  })

  it('puts the header before normal content when no leading Word metadata exists', () => {
    const mjml = buildMjml(documentWith({ type: 'paragraph', runs: [{ text: 'Olá' }] }), settings)
    expect(mjml.indexOf('bgcolor="#000000"')).toBeLessThan(mjml.indexOf('Olá'))
    expect(mjml).toContain('<mj-body width="600px"')
  })

  it('omits the header completely when disabled and keeps the original content section', () => {
    const mjml = buildMjml(documentWith({ type: 'paragraph', runs: [{ text: 'Olá' }] }), {
      ...settings,
      header: { ...settings.header, enabled: false },
    })
    expect(mjml).not.toContain('<mj-table')
    expect(mjml).not.toContain('bgcolor="#000000"')
    expect(mjml).toContain('<mj-section padding="24px 16px"')
  })

  it('uses configured color, height and vertical spacing without changing global width', () => {
    const mjml = buildMjml(documentWith({ type: 'paragraph', runs: [{ text: 'Olá' }] }), {
      ...settings,
      header: {
        enabled: true,
        backgroundColor: '#123456',
        height: 144,
        marginTop: 12,
        marginBottom: 18,
      },
    })
    expect(mjml).toContain('<mj-body width="600px"')
    expect(mjml).toContain('bgcolor="#123456" height="144"')
    expect(mjml).toContain('background-color:#123456;height:144px')
    expect(mjml).toContain('<mj-spacer height="12px" padding="0" />')
    expect(mjml).toContain('<mj-spacer height="18px" padding="0" />')
    expect(mjml).toContain('<mj-table padding="0" width="100%"')
  })
})

describe('renderEmail', () => {
  it('exports a single full-width 100 px header with no white gutter at the default settings', async () => {
    const result = await renderEmail(documentWith({ type: 'paragraph', runs: [{ text: 'Olá' }] }), settings)
    expect(result.errors).toEqual([])
    const parsed = new DOMParser().parseFromString(result.html, 'text/html')
    const cell = parsed.querySelector('td[bgcolor="#000000"]')
    expect(cell).not.toBeNull()
    expect(cell?.getAttribute('height')).toBe('100')
    expect(cell?.getAttribute('style')).toContain('height:100px')
    expect(cell?.getAttribute('style')).toContain('mso-line-height-rule:exactly')
    expect(cell?.closest('table')?.getAttribute('width')).toBe('100%')
    expect(cell?.closest('table')?.parentElement?.getAttribute('style')).toContain('padding:0')
    expect(parsed.querySelectorAll('td[bgcolor="#000000"]')).toHaveLength(1)
    expect(result.html).toContain('vertical-align:top;width:568px;')
  })

  it('exports edited header dimensions, color and spacing, and omits it when disabled', async () => {
    const document = documentWith({ type: 'paragraph', runs: [{ text: 'Corpo' }] })
    const customized = await renderEmail(document, {
      ...settings,
      header: { enabled: true, backgroundColor: '#123456', height: 144, marginTop: 12, marginBottom: 18 },
    })
    const parsed = new DOMParser().parseFromString(customized.html, 'text/html')
    const cell = parsed.querySelector('td[bgcolor="#123456"]')
    expect(customized.errors).toEqual([])
    expect(cell?.getAttribute('height')).toBe('144')
    expect(cell?.getAttribute('style')).toContain('background-color:#123456;height:144px')
    expect(customized.html).toContain('height:12px')
    expect(customized.html).toContain('height:18px')

    const disabled = await renderEmail(document, {
      ...settings,
      header: { ...settings.header, enabled: false },
    })
    expect(disabled.errors).toEqual([])
    expect(disabled.html).not.toContain('bgcolor="#000000"')
    expect(disabled.html).not.toContain('bgcolor="#123456"')
    expect(disabled.html).toContain('Corpo')
  })

  it('keeps a Word-like subject, fine line and formatted copy in order around the header', async () => {
    const parsed = parseMammothHtml(
      '<p>Assunto: <strong>Oferta</strong></p><p>Linha fina: Detalhes</p><p>Olá <em>cliente</em>.</p>',
    )
    const result = await renderEmail({ ...parsed, sourceName: 'copy.docx' }, settings)
    const header = result.html.indexOf('bgcolor="#000000"')
    expect(result.errors).toEqual([])
    expect(result.html.indexOf('Assunto:')).toBeLessThan(header)
    expect(result.html.indexOf('Linha fina:')).toBeLessThan(header)
    expect(header).toBeLessThan(result.html.indexOf('Olá'))
    expect(result.html).toContain('<strong>Oferta</strong>')
    expect(result.html).toContain('<em>cliente</em>')
  })

  it('splits Word soft line breaks after the fine line without losing runs or links', async () => {
    const parsed = parseMammothHtml(
      '<p>Assunto: <strong>Oferta</strong><br />Linha fina: <em>Detalhes</em><br />Olá <a href="https://example.org/oferta">cliente</a></p>',
    )
    const result = await renderEmail({ ...parsed, sourceName: 'linhas.docx' }, settings)
    const header = result.html.indexOf('bgcolor="#000000"')
    expect(result.errors).toEqual([])
    expect(result.html.indexOf('Assunto:')).toBeLessThan(header)
    expect(result.html.indexOf('Linha fina:')).toBeLessThan(header)
    expect(header).toBeLessThan(result.html.indexOf('Olá'))
    expect(result.html).toContain('<strong>Oferta</strong>')
    expect(result.html).toContain('<em>Detalhes</em>')
    expect(result.html).toContain('href="https://example.org/oferta"')
    expect(result.html).toContain('>cliente</a>')
    expect(result.html.match(/bgcolor="#000000"/g)).toHaveLength(1)
  })

  it('keeps the header outside an AJO condition even when the Word starts with that condition', async () => {
    const document = documentWith(
      { type: 'condition-start', runs: [], audience: 'BASE_ATIVO' },
      { type: 'paragraph', runs: [{ text: 'Oferta condicional' }] },
      { type: 'condition-end', runs: [], audience: 'BASE_ATIVO' },
    )
    const first = await renderEmail(document, settings)
    const second = await renderEmail(document, settings)
    const opening = '{%#if profile.segmentMembership.get("ups").get("BASE_ATIVO").status = "existing"%}'
    expect(first.errors).toEqual([])
    expect(first.html.indexOf('bgcolor="#000000"')).toBeLessThan(first.html.indexOf(opening))
    expect(first.html.indexOf(opening)).toBeLessThan(first.html.indexOf('Oferta condicional'))
    expect(first.html.match(/bgcolor="#000000"/g)).toHaveLength(1)
    expect(second.html.match(/bgcolor="#000000"/g)).toHaveLength(1)
    expect(second.html).toBe(first.html)
  })

  it('exports compiled email HTML with the original copy and no unsafe HTML', async () => {
    const document = documentWith({
      type: 'paragraph',
      runs: [
        { text: 'Olá ' },
        { text: 'cliente', bold: true },
        { text: ' <img src=x onerror=alert(1)>' },
      ],
    })

    const result = await renderEmail(document, settings)
    expect(result.errors).toEqual([])
    expect(result.html).toContain('Olá')
    expect(result.html).toContain('<strong>cliente</strong>')
    expect(result.html).toContain('&lt;img src=x onerror=alert(1)&gt;')
    expect(result.html).not.toContain('<img src=x onerror=alert(1)>')
  })

  it('inserts every original table button after MJML compilation in Word order', async () => {
    const document = documentWith(
      { type: 'paragraph', runs: [{ text: 'Antes ' }, { text: 'forte', bold: true }] },
      { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
      { type: 'paragraph', runs: [{ text: 'Entre ', italic: true }, { text: 'os botões' }] },
      { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-2' },
      { type: 'paragraph', runs: [{ text: 'Depois' }] },
    )
    const result = await renderEmail(document, settings)
    expect(result.errors).toEqual([])
    expect(result.html.split(ACESSAR_HTML)).toHaveLength(3)
    const first = result.html.indexOf(ACESSAR_HTML)
    const second = result.html.indexOf(ACESSAR_HTML, first + 1)
    expect(result.html.indexOf('<strong>forte</strong>')).toBeLessThan(first)
    expect(first).toBeLessThan(result.html.indexOf('<em>Entre </em>'))
    expect(result.html.indexOf('<em>Entre </em>')).toBeLessThan(second)
    expect(second).toBeLessThan(result.html.indexOf('Depois'))
    expect(result.html).not.toContain('%acessar%')
    expect(result.html).not.toContain('__ATELIE_COMPONENT_SLOT_')
  })

  it('edits only the selected button text and URL while preserving the original table and styles', async () => {
    const document = documentWith(
      { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
      { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-2' },
    )
    const edited = renderCustomComponent('acessar', { text: 'Abrir <oferta>', url: 'https://example.com/?a=1&b=2' })
    const result = await renderEmail(document, settings, {
      'acessar-1': { text: 'Abrir <oferta>', url: 'https://example.com/?a=1&b=2' },
    })
    expect(result.errors).toEqual([])
    expect(result.html).toContain(edited)
    expect(result.html).toContain(ACESSAR_HTML)
    expect(result.html).not.toContain('<oferta>')
    expect(edited).toContain('href="https://example.com/?a=1&amp;b=2"')
    expect(edited).toContain('Abrir &lt;oferta&gt;')
    expect(edited).toContain('bgcolor="#FD5E04"')
    expect(edited).toContain('padding:20px;border-radius:7px')
  })

  it('blocks preview and export when a button URL is invalid', async () => {
    const result = await renderEmail(
      documentWith({ type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' }),
      settings,
      { 'acessar-1': { text: 'Teste', url: 'javascript:alert(1)' } },
    )
    expect(result.html).toBe('')
    expect(result.errors.join(' ')).toMatch(/URL HTTP ou HTTPS válida/)
  })

  it('exports a Word-linked %texto% button with the original table markup in the correct position', async () => {
    const url = 'https://example.org/oferta?a=1&b=2'
    const button = {
      type: 'component' as const, runs: [] as [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url },
    }
    const document = documentWith(
      { type: 'paragraph', runs: [{ text: 'Antes' }] },
      button,
      { type: 'paragraph', runs: [{ text: 'Depois' }] },
    )

    const result = await renderEmail(document, settings)
    const expectedButton = renderCustomComponent('word-button', button.initial)
    const position = result.html.indexOf(expectedButton)

    expect(result.errors).toEqual([])
    expect(position).toBeGreaterThan(-1)
    expect(result.html.indexOf('Antes')).toBeLessThan(position)
    expect(position).toBeLessThan(result.html.indexOf('Depois'))
    expect(result.html).not.toContain('%BAIXAR AGORA%')
    expect(expectedButton).toContain('bgcolor="#FD5E04"')
    expect(expectedButton).toContain('padding:20px;border-radius:7px')
    expect(expectedButton).toContain('class="innertd buttonblock"')
    expect(expectedButton).toContain('class="buttonstyles"')
    expect(expectedButton).toContain('data-nl-type="externalLink"')
    expect(expectedButton).toContain('href="https://example.org/oferta?a=1&amp;b=2"')
    expect(expectedButton).toContain('>BAIXAR AGORA</a>')
  })

  it('keeps each Word button independent when its text and URL are edited', async () => {
    const document = documentWith(
      {
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'BAIXAR AGORA', url: 'https://example.org/oferta' },
      },
      { type: 'paragraph', runs: [{ text: 'Entre' }] },
      {
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-2',
        initial: { text: 'SAIBA MAIS', url: 'https://example.org/mais' },
      },
    )
    const edited = renderCustomComponent('word-button', { text: 'Abrir <agora>', url: 'https://example.org/novo?a=1&b=2' })
    const unchanged = renderCustomComponent('word-button', { text: 'SAIBA MAIS', url: 'https://example.org/mais' })

    const result = await renderEmail(document, settings, {
      'word-button-1': { text: 'Abrir <agora>', url: 'https://example.org/novo?a=1&b=2' },
    })

    expect(result.errors).toEqual([])
    expect(result.html).toContain(edited)
    expect(result.html).toContain(unchanged)
    expect(result.html.indexOf(edited)).toBeLessThan(result.html.indexOf('Entre'))
    expect(result.html.indexOf('Entre')).toBeLessThan(result.html.indexOf(unchanged))
    expect(edited).toContain('>Abrir &lt;agora&gt;</a>')
    expect(edited).not.toContain('<agora>')
    expect(edited).toBe(
      ACESSAR_HTML
        .replace(/href="[^"]*"/, 'href="https://example.org/novo?a=1&amp;b=2"')
        .replace('>Teste doc</a>', '>Abrir &lt;agora&gt;</a>'),
    )
  })

  it('exports a Word button with # when its marker has no hyperlink', async () => {
    const document = documentWith({
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url: '#' },
    })

    const fallback = await renderEmail(document, settings)
    expect(fallback.errors).toEqual([])
    expect(fallback.html).toContain('href="#"')
    expect(fallback.html).toContain('>BAIXAR AGORA</a>')

    const configured = await renderEmail(document, settings, {
      'word-button-1': { url: 'https://example.org/download' },
    })
    expect(configured.errors).toEqual([])
    expect(configured.html).toContain('href="https://example.org/download"')
    expect(configured.html).toContain('>BAIXAR AGORA</a>')
  })

  it('exports # after a button URL is cleared in the editor', async () => {
    const result = await renderEmail(
      documentWith({
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'BAIXAR AGORA', url: 'https://example.org/download' },
      }),
      settings,
      { 'word-button-1': { url: '' } },
    )

    expect(result.errors).toEqual([])
    expect(result.html).toContain('href="#"')
    expect(result.html).not.toContain('href="https://example.org/download"')
    expect(result.html).toContain('>BAIXAR AGORA</a>')
  })

  it('wraps complete MJML rows and the original button HTML in a literal AJO condition', async () => {
    const opening = '{%#if profile.segmentMembership.get("ups").get("CLIENTES_VIP").status = "existing"%}'
    const document = documentWith(
      { type: 'paragraph', runs: [{ text: 'Antes' }] },
      { type: 'condition-start', runs: [], audience: 'CLIENTES_VIP' },
      { type: 'paragraph', runs: [{ text: 'Oferta ' }, { text: 'especial', bold: true }] },
      { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
      { type: 'condition-end', runs: [], audience: 'CLIENTES_VIP' },
      { type: 'paragraph', runs: [{ text: 'Depois' }] },
    )
    const result = await renderEmail(document, settings)
    expect(result.errors).toEqual([])
    const start = result.html.indexOf(opening)
    const end = result.html.indexOf('{%/if%}')
    expect(start).toBeGreaterThan(result.html.indexOf('Antes'))
    expect(end).toBeGreaterThan(start)
    expect(result.html.indexOf('Depois')).toBeGreaterThan(end)
    const protectedHtml = result.html.slice(start, end)
    expect(protectedHtml).toContain('<strong>especial</strong>')
    expect(protectedHtml).toContain(ACESSAR_HTML)
    expect(result.html).not.toContain('__ATELIE_CONDITION_SLOT_')
    expect(result.html).not.toContain('##CLIENTES_VIP##')
    // Raw slots sit between complete table rows after MJML compilation.
    expect(result.html.slice(start - 30, start)).toMatch(/<\/tr>\s*$/)
    expect(result.html.slice(end + '{%/if%}'.length, end + '{%/if%}'.length + 50)).toMatch(/^\s*<tr>/)
  })

  it('exports independent conditions for different and repeated audiences', async () => {
    const document = documentWith(
      { type: 'condition-start', runs: [], audience: 'BASE_ATIVO' },
      { type: 'paragraph', runs: [{ text: 'Primeiro' }] },
      { type: 'condition-end', runs: [], audience: 'BASE_ATIVO' },
      { type: 'paragraph', runs: [{ text: 'Fora' }] },
      { type: 'condition-start', runs: [], audience: 'CLIENTES_VIP' },
      { type: 'paragraph', runs: [{ text: 'Segundo' }] },
      { type: 'condition-end', runs: [], audience: 'CLIENTES_VIP' },
      { type: 'condition-start', runs: [], audience: 'BASE_ATIVO' },
      { type: 'paragraph', runs: [{ text: 'Terceiro' }] },
      { type: 'condition-end', runs: [], audience: 'BASE_ATIVO' },
    )
    const result = await renderEmail(document, settings)
    expect(result.errors).toEqual([])
    expect(result.html.match(/get\("BASE_ATIVO"\)/g)).toHaveLength(2)
    expect(result.html.match(/get\("CLIENTES_VIP"\)/g)).toHaveLength(1)
    expect(result.html.match(/\{%\/if%\}/g)).toHaveLength(3)
    expect(result.html.indexOf('Primeiro')).toBeLessThan(result.html.indexOf('Fora'))
    expect(result.html.indexOf('Fora')).toBeLessThan(result.html.indexOf('Segundo'))
    expect(result.html.indexOf('Segundo')).toBeLessThan(result.html.indexOf('Terceiro'))
  })

  it('blocks malformed conditions before generating exportable HTML', async () => {
    const invalid = documentWith(
      { type: 'condition-start', runs: [], audience: 'BASE_ATIVO' },
      { type: 'paragraph', runs: [{ text: 'Texto' }] },
      { type: 'condition-end', runs: [], audience: 'CLIENTES_VIP' },
    )
    const result = await renderEmail(invalid, settings)
    expect(result.html).toBe('')
    expect(result.errors.join(' ')).toMatch(/não corresponde/)
    const inline = await renderEmail(documentWith({ type: 'paragraph', runs: [{ text: 'Leia ##BASE_ATIVO## aqui' }] }), settings)
    expect(inline.html).toBe('')
    expect(inline.errors.join(' ')).toMatch(/sozinho em um parágrafo/)
  })
})
