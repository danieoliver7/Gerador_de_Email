import { describe, expect, it } from 'vitest'
import { parseMammothHtml } from './document'
import type { ContentBlock, ExtractedDocument } from './document'
import { ACESSAR_HTML, registerCustomComponent, renderCustomComponent } from './custom-components'
import { renderEmail } from './email'
import { DEFAULT_SETTINGS } from './settings'

const settings = { ...DEFAULT_SETTINGS, header: { ...DEFAULT_SETTINGS.header } }
const paragraph = (text: string): ContentBlock => ({ type: 'paragraph', runs: [{ text }] })
const button = (number: number): ContentBlock => ({
  type: 'component', runs: [], componentId: 'word-button', instanceId: `word-button-${number}`,
  initial: { text: `Oferta ${number}`, url: `https://example.org/oferta-${number}` },
})
const start = (audience = 'BASE_ATIVO'): ContentBlock => ({ type: 'condition-start', runs: [], audience })
const end = (audience = 'BASE_ATIVO'): ContentBlock => ({ type: 'condition-end', runs: [], audience })
const documentWith = (...blocks: ContentBlock[]): ExtractedDocument => ({ blocks, sourceName: 'copy.docx', warnings: [] })
const opening = (name: string) => `<!-- ${name} - INÍCIO -->`
const closing = (name: string) => `<!-- ${name} - FIM -->`
const ORGANIZATION_COMMENT = /<!--\s*((?:CABEÇALHO|PARTE (?:\d+|FINAL)|BOTÃO \d+)) - (INÍCIO|FIM)\s*-->/gu

function sectionNames(html: string): string[] {
  return [...html.matchAll(ORGANIZATION_COMMENT)]
    .filter((match) => match[2] === 'INÍCIO')
    .map((match) => match[1])
}

function expectSafeConditionBoundaries(html: string): void {
  let audienceDepth = 0
  let active: { name: string; depth: number } | null = null
  const tokens = /<!--\s*((?:CABEÇALHO|PARTE (?:\d+|FINAL)|BOTÃO \d+)) - (INÍCIO|FIM)\s*-->|\{%#if[^]*?%\}|\{%\/if%\}/gu
  for (const match of html.matchAll(tokens)) {
    if (match[0].startsWith('{%')) {
      // A logical part may resume later, but each annotated segment closes
      // before an AJO boundary, so removing an audience cannot unpair comments.
      expect(active, `section crosses ${match[0]}`).toBeNull()
      audienceDepth += match[0].startsWith('{%#if') ? 1 : -1
      expect(audienceDepth).toBeGreaterThanOrEqual(0)
      continue
    }
    if (match[2] === 'INÍCIO') {
      expect(active).toBeNull()
      active = { name: match[1], depth: audienceDepth }
    } else {
      expect(active).toEqual({ name: match[1], depth: audienceDepth })
      active = null
    }
  }
  expect(active).toBeNull()
  expect(audienceDepth).toBe(0)
}

function expectCompleteRowComments(html: string): void {
  const parsed = new DOMParser().parseFromString(html, 'text/html')
  const walker = parsed.createTreeWalker(parsed, NodeFilter.SHOW_COMMENT)
  let node = walker.nextNode()
  let count = 0
  while (node) {
    if (/^(CABEÇALHO|PARTE (?:\d+|FINAL)|BOTÃO \d+) - (INÍCIO|FIM)$/u.test(node.textContent?.trim() ?? '')) {
      count += 1
      expect(['TD', 'TR', 'A', 'P', 'V:TEXTBOX']).not.toContain(node.parentElement?.tagName)
      if (node.parentElement?.tagName === 'TBODY') {
        const previous = node.previousSibling
        const next = node.nextSibling
        if (previous?.nodeType === Node.ELEMENT_NODE) expect((previous as Element).tagName).toBe('TR')
        if (next?.nodeType === Node.ELEMENT_NODE) expect((next as Element).tagName).toBe('TR')
      }
    }
    node = walker.nextNode()
  }
  expect(count).toBeGreaterThan(0)
}

describe('organized email export', () => {
  it.each([0, 1, 2, 4, 6])('numbers the complete %i-button document and its content parts', async (count) => {
    const blocks: ContentBlock[] = []
    for (let index = 1; index <= count; index += 1) blocks.push(paragraph(`Conteúdo ${index}`), button(index))
    if (count === 0) blocks.push(paragraph('Corpo sem botões'))
    else blocks.push(paragraph('Assinatura final'))
    const result = await renderEmail(documentWith(...blocks), settings)
    expect(result.errors).toEqual([])
    const names = ['CABEÇALHO']
    if (count === 0) names.push('PARTE 1')
    else {
      for (let index = 1; index <= count; index += 1) names.push(`PARTE ${index}`, `BOTÃO ${index}`)
      names.push('PARTE FINAL')
    }
    expect(sectionNames(result.html)).toEqual(names)
    for (const name of names) {
      expect(result.html.split(opening(name))).toHaveLength(2)
      expect(result.html.split(closing(name))).toHaveLength(2)
    }
    expectSafeConditionBoundaries(result.html)
    expectCompleteRowComments(result.html)
  })

  it('does not insert empty parts before adjacent buttons or after the last button', async () => {
    const result = await renderEmail(documentWith(button(1), button(2)), settings)
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html)).toEqual(['CABEÇALHO', 'BOTÃO 1', 'BOTÃO 2'])
    expect(result.html).not.toContain('PARTE FINAL')
    expect(result.html).not.toContain('PARTE 1')
    expect(result.html).not.toContain('PARTE 2')
  })

  it('only creates a final part when there is trailing content', async () => {
    const withoutFinal = await renderEmail(documentWith(paragraph('Introdução'), button(1)), settings)
    const withFinal = await renderEmail(documentWith(paragraph('Introdução'), button(1), paragraph('Equipe InfoMoney')), settings)
    expect(sectionNames(withoutFinal.html)).toEqual(['CABEÇALHO', 'PARTE 1', 'BOTÃO 1'])
    expect(sectionNames(withFinal.html)).toEqual(['CABEÇALHO', 'PARTE 1', 'BOTÃO 1', 'PARTE FINAL'])
    expect(withFinal.html.indexOf(opening('PARTE FINAL'))).toBeLessThan(withFinal.html.indexOf('Equipe InfoMoney'))
    expect(withFinal.html.indexOf('Equipe InfoMoney')).toBeLessThan(withFinal.html.indexOf(closing('PARTE FINAL')))
  })

  it('places the unique header after subject and preheader while preserving imported inline formatting', async () => {
    const parsed = parseMammothHtml(
      '<p>Assunto: <strong>Oferta</strong></p><p>Pré-header: <em>Detalhes</em></p><p>Olá <a href="https://example.org/oferta">cliente</a>.</p>',
    )
    const original = structuredClone(parsed.blocks)
    const result = await renderEmail({ ...parsed, sourceName: 'metadata.docx' }, settings)
    const header = result.html.indexOf('bgcolor="#000000"')
    expect(result.errors).toEqual([])
    expect(result.html.indexOf('Assunto:')).toBeLessThan(header)
    expect(result.html.indexOf('Pré-header:')).toBeLessThan(header)
    expect(header).toBeLessThan(result.html.indexOf('Olá'))
    expect(result.html.split('Assunto:')).toHaveLength(2)
    expect(result.html.split('Pré-header:')).toHaveLength(2)
    expect(result.html.split('bgcolor="#000000"')).toHaveLength(2)
    expect(result.html).toContain('<strong>Oferta</strong>')
    expect(result.html).toContain('<em>Detalhes</em>')
    expect(result.html).toContain('href="https://example.org/oferta"')
    expect(parsed.blocks).toEqual(original)
    const dom = new DOMParser().parseFromString(result.html, 'text/html')
    const cell = dom.querySelector('td[bgcolor="#000000"]')
    expect(cell?.getAttribute('height')).toBe('100')
    expect(cell?.closest('table')?.getAttribute('width')).toBe('100%')
    expect(cell?.closest('table')?.parentElement?.getAttribute('style')).toContain('padding:0')
  })

  it('does not manufacture subject or preheader metadata for a body-only Word document', async () => {
    const result = await renderEmail(documentWith(paragraph('Olá, cliente!')), settings)
    expect(result.errors).toEqual([])
    expect(result.html).not.toContain('Assunto:')
    expect(result.html).not.toContain('Pré-header:')
    expect(result.html.indexOf('bgcolor="#000000"')).toBeLessThan(result.html.indexOf('Olá, cliente!'))
  })

  it('omits header comments and markup when the existing setting disables it', async () => {
    const result = await renderEmail(documentWith(paragraph('Introdução'), button(1)), {
      ...settings, header: { ...settings.header, enabled: false },
    })
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html)).toEqual(['PARTE 1', 'BOTÃO 1'])
    expect(result.html).not.toContain('bgcolor="#000000"')
  })

  it('counts a conditional first button without inventing a preceding empty part', async () => {
    const result = await renderEmail(documentWith(start(), button(1), end(), paragraph('Entre'), button(2)), settings)
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html)).toEqual(['CABEÇALHO', 'BOTÃO 1', 'PARTE 2', 'BOTÃO 2'])
    expect(result.html.indexOf('{%#if')).toBeLessThan(result.html.indexOf(opening('BOTÃO 1')))
    expect(result.html.indexOf(closing('BOTÃO 1'))).toBeLessThan(result.html.indexOf('{%/if%}'))
    expectSafeConditionBoundaries(result.html)
  })

  it('counts buttons in different audiences between common buttons, preserving the exact Adobe expressions', async () => {
    const document = documentWith(
      paragraph('Intro'), button(1), paragraph('Entre 1 e 2'),
      start(), paragraph('Oferta ativa'), button(2), end(),
      paragraph('Entre 2 e 3'), start('CLIENTES_VIP'), button(3), end('CLIENTES_VIP'),
      paragraph('Entre 3 e 4'), button(4), paragraph('Assinatura'),
    )
    const result = await renderEmail(document, settings)
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html).filter((name) => name.startsWith('BOTÃO')))
      .toEqual(['BOTÃO 1', 'BOTÃO 2', 'BOTÃO 3', 'BOTÃO 4'])
    expect(result.html).toContain('{%#if profile.segmentMembership.get("ups").get("BASE_ATIVO").status = "existing"%}')
    expect(result.html).toContain('{%#if profile.segmentMembership.get("ups").get("CLIENTES_VIP").status = "existing"%}')
    expect(result.html.match(/\{%\/if%\}/gu)).toHaveLength(2)
    expectSafeConditionBoundaries(result.html)
    expectCompleteRowComments(result.html)
  })

  it('retains one condition around multiple distinct buttons and their original text/list scope', async () => {
    const result = await renderEmail(documentWith(
      paragraph('Antes'), start(), paragraph('Exclusivo'), button(1),
      { type: 'list-item', runs: [{ text: 'Benefício exclusivo' }], ordered: false, level: 1 },
      button(2), paragraph('Conclusão exclusiva'), end(), paragraph('Depois'),
    ), settings)
    expect(result.errors).toEqual([])
    expect(result.html.match(/\{%#if/gu)).toHaveLength(1)
    expect(result.html.match(/\{%\/if%\}/gu)).toHaveLength(1)
    const protectedHtml = result.html.slice(result.html.indexOf('{%#if'), result.html.indexOf('{%/if%}'))
    for (const text of ['Exclusivo', 'Benefício exclusivo', 'Oferta 1', 'Oferta 2', 'Conclusão exclusiva']) {
      expect(protectedHtml).toContain(text)
    }
    expect(protectedHtml).not.toContain('Antes')
    expect(protectedHtml).not.toContain('Depois')
    expectSafeConditionBoundaries(result.html)
  })

  it('splits the same logical part safely at both condition boundaries without changing its number', async () => {
    const result = await renderEmail(documentWith(
      paragraph('Introdução comum'), start(), paragraph('Introdução protegida'), end(),
      paragraph('Continuação comum'), button(1), paragraph('Assinatura'),
    ), settings)
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html)).toEqual(['CABEÇALHO', 'PARTE 1', 'PARTE 1', 'PARTE 1', 'BOTÃO 1', 'PARTE FINAL'])
    expectSafeConditionBoundaries(result.html)
  })

  it('keeps full-document numbering when a non-member preview hides conditional buttons', async () => {
    const document = documentWith(
      paragraph('Introdução'), button(1), start(), paragraph('Texto exclusivo'), button(2), end(),
      start('CLIENTES_VIP'), button(3), end('CLIENTES_VIP'), paragraph('Retomada'), button(4), paragraph('Final'),
    )
    const original = structuredClone(document)
    const exported = await renderEmail(document, settings)
    const member = await renderEmail(document, settings, {}, { audienceSimulation: 'member' })
    const nonMember = await renderEmail(document, settings, {}, { audienceSimulation: 'non-member' })
    expect(exported.errors).toEqual([])
    expect(member.errors).toEqual([])
    expect(nonMember.errors).toEqual([])
    expect(sectionNames(exported.html).filter((name) => name.startsWith('BOTÃO')))
      .toEqual(['BOTÃO 1', 'BOTÃO 2', 'BOTÃO 3', 'BOTÃO 4'])
    expect(sectionNames(member.html).filter((name) => name.startsWith('BOTÃO')))
      .toEqual(['BOTÃO 1', 'BOTÃO 2', 'BOTÃO 3', 'BOTÃO 4'])
    expect(sectionNames(nonMember.html).filter((name) => name.startsWith('BOTÃO')))
      .toEqual(['BOTÃO 1', 'BOTÃO 4'])
    expect(nonMember.html).toContain(opening('PARTE 4'))
    expect(nonMember.html).not.toContain('Texto exclusivo')
    expect(nonMember.html).not.toContain('Oferta 2')
    expect(nonMember.html).not.toContain('Oferta 3')
    expect(member.html).toContain('Texto exclusivo')
    expect(member.html).not.toContain('{%')
    expect(nonMember.html).not.toContain('{%')
    expect(document).toEqual(original)
    expectSafeConditionBoundaries(member.html)
    expectSafeConditionBoundaries(nonMember.html)
  })

  it('counts only button components rather than ordinary hyperlinks in imported text', async () => {
    const parsed = parseMammothHtml(
      '<p>Leia <a href="https://example.org/noticia">a notícia</a>.</p>' +
      '<p><a href="https://example.org/oferta">%SAIBA MAIS%</a></p>' +
      '<p><a href="https://example.org/rodape">Rodapé</a></p>',
    )
    const result = await renderEmail({ ...parsed, sourceName: 'links.docx' }, settings)
    expect(result.errors).toEqual([])
    expect(sectionNames(result.html).filter((name) => name.startsWith('BOTÃO'))).toEqual(['BOTÃO 1'])
    expect(result.html).toContain('href="https://example.org/noticia"')
    expect(result.html).toContain('href="https://example.org/rodape"')
    expect(result.html).toContain('data-nl-type="externalLink"')
  })

  it('preserves the entire original button and only wraps it with organization comments', async () => {
    const result = await renderEmail(documentWith({ type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' }), settings)
    const section = result.html.slice(result.html.indexOf(opening('BOTÃO 1')), result.html.indexOf(closing('BOTÃO 1')))
    expect(result.errors).toEqual([])
    expect(section).toContain(ACESSAR_HTML)
    expect(result.html.split(ACESSAR_HTML)).toHaveLength(2)
    expectCompleteRowComments(result.html)
  })

  it('preserves trusted Outlook VML and conditional comments byte for byte', async () => {
    const outlook = '<!--[if mso]><v:roundrect xmlns:v="urn:schemas-microsoft-com:vml" href="https://example.org/oferta" style="height:60px;v-text-anchor:middle;width:220px;" arcsize="12%" stroke="f" fillcolor="#FD5E04"><v:textbox inset="0,0,0,0"><center style="color:#ffffff;font-family:Arial;font-size:16px;">Oferta Outlook</center></v:textbox></v:roundrect><![endif]-->'
    const html = outlook + '<!--[if !mso]><!-->' + ACESSAR_HTML + '<!--<![endif]-->'
    registerCustomComponent({ id: 'qa-outlook-button', name: 'Botão Outlook para teste', marker: null, html })
    const result = await renderEmail(documentWith({ type: 'component', runs: [], componentId: 'qa-outlook-button', instanceId: 'qa-outlook-button-1' }), settings)
    expect(result.errors).toEqual([])
    expect(result.html).toContain(html)
    expect(result.html.split(outlook)).toHaveLength(2)
    expect(sectionNames(result.html)).toEqual(['CABEÇALHO', 'BOTÃO 1'])
    expectCompleteRowComments(result.html)
  })

  it('preserves independent overrides including Adobe attributes and Outlook surrounding markup', async () => {
    const document = documentWith(button(1), paragraph('Entre'), button(2))
    const overrides = { 'word-button-1': { text: 'Abrir <agora>', url: 'https://example.org/novo?a=1&b=2' } }
    const result = await renderEmail(document, settings, overrides)
    expect(result.errors).toEqual([])
    expect(result.html).toContain(renderCustomComponent('word-button', overrides['word-button-1']))
    expect(result.html).toContain(renderCustomComponent('word-button', (button(2) as Extract<ContentBlock, { type: 'component' }>).initial))
    expect(result.html).toContain('data-nl-type="externalLink"')
    expect(result.html).toContain('mso-line-height-rule:exactly')
  })

  it('regenerates deterministically and renumbers after a structural edit without mutating the document', async () => {
    const document = documentWith(paragraph('Introdução'), button(1), paragraph('Meio'), button(3), paragraph('Final'))
    const original = structuredClone(document)
    const first = await renderEmail(document, settings)
    const second = await renderEmail(document, settings)
    expect(first.errors).toEqual([])
    expect(second.html).toBe(first.html)
    expect(sectionNames(first.html).filter((name) => name.startsWith('BOTÃO'))).toEqual(['BOTÃO 1', 'BOTÃO 2'])
    expect(first.html.split(opening('BOTÃO 1'))).toHaveLength(2)
    expect(first.html.split(opening('BOTÃO 2'))).toHaveLength(2)
    const changed = await renderEmail(documentWith(paragraph('Introdução'), button(3), paragraph('Final')), settings)
    expect(sectionNames(changed.html).filter((name) => name.startsWith('BOTÃO'))).toEqual(['BOTÃO 1'])
    expect(document).toEqual(original)
  })

  it('still rejects malformed Adobe conditions before generating exportable HTML', async () => {
    const result = await renderEmail(documentWith(start(), button(1), end('CLIENTES_VIP')), settings)
    expect(result.html).toBe('')
    expect(result.errors.join(' ')).toContain('não corresponde')
  })
})
