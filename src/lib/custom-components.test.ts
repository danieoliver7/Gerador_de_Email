import { describe, expect, it } from 'vitest'
import type { ContentBlock } from './document'
import {
  ACESSAR_HTML,
  getComponentDefinition,
  registerCustomComponent,
  renderCustomComponent,
  renderCustomComponentDraft,
  resolveComponentMarkers,
} from './custom-components'

function paragraph(...texts: string[]): ContentBlock {
  return { type: 'paragraph', runs: texts.map((text) => ({ text })) }
}

describe('custom component markers', () => {
  it('replaces one standalone marker with a component block', () => {
    const result = resolveComponentMarkers([paragraph('%acessar%')])
    expect(result).toEqual({
      blocks: [{ type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' }],
      warnings: [],
    })
  })

  it('keeps multiple marker instances in their original positions', () => {
    const result = resolveComponentMarkers([
      paragraph('Antes'),
      paragraph('%acessar%'),
      paragraph('Entre'),
      paragraph('%acessar%'),
      paragraph('Depois'),
    ])
    expect(result.blocks.map((block) => block.type === 'component' ? block.instanceId : block.runs.map((run) => run.text).join('')))
      .toEqual(['Antes', 'acessar-1', 'Entre', 'acessar-2', 'Depois'])
  })

  it('recognizes a marker split across formatted Word runs', () => {
    const result = resolveComponentMarkers([{
      type: 'paragraph',
      runs: [{ text: '  %ac' }, { text: 'es', bold: true }, { text: 'sar%', italic: true }, { text: '  ' }],
    }])
    expect(result.blocks).toEqual([{ type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' }])
  })

  it('preserves copy and formatting on either side of a standalone marker', () => {
    const before: ContentBlock = { type: 'paragraph', runs: [{ text: 'Olá ' }, { text: 'cliente', bold: true }] }
    const after: ContentBlock = { type: 'paragraph', runs: [{ text: 'Veja mais', italic: true, href: 'https://example.org' }] }
    const result = resolveComponentMarkers([before, paragraph('%acessar%'), after])
    expect(result.blocks[0]).toBe(before)
    expect(result.blocks[2]).toBe(after)
    expect(result.blocks[1]).toMatchObject({ type: 'component', componentId: 'acessar' })
  })

  it('warns about inline markers and retains every run unchanged', () => {
    const inline: ContentBlock = { type: 'paragraph', runs: [{ text: 'Clique ' }, { text: '%ac', bold: true }, { text: 'essar%' }, { text: ' agora' }] }
    const result = resolveComponentMarkers([inline])
    expect(result.blocks).toEqual([inline])
    expect(result.warnings).toHaveLength(1)
    expect(result.warnings[0]).toContain('%acessar%')
  })

  it('turns any standalone %texto% into a button with its Word hyperlink and label', () => {
    const result = resolveComponentMarkers([{
      type: 'paragraph',
      runs: [
        { text: '%' },
        { text: 'BAIXAR AGORA', href: 'https://example.org/planilha?ref=word', bold: true },
        { text: '%' },
      ],
    }])
    expect(result).toEqual({
      blocks: [{
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'BAIXAR AGORA', url: 'https://example.org/planilha?ref=word' },
      }],
      warnings: [],
    })
  })

  it('recognizes a generic marker split across differently formatted Word runs', () => {
    const result = resolveComponentMarkers([{
      type: 'paragraph',
      runs: [
        { text: '  %BA', italic: true },
        { text: 'IXAR ', href: 'https://example.org/arquivo' },
        { text: 'AGORA', bold: true, href: 'https://example.org/arquivo' },
        { text: '%  ' },
      ],
    }])
    expect(result.blocks[0]).toEqual({
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url: 'https://example.org/arquivo' },
    })
  })

  it('splits one Word paragraph around a marker line and preserves surrounding run formatting', () => {
    const result = resolveComponentMarkers([{
      type: 'paragraph',
      runs: [
        { text: 'Antes\n', bold: true },
        { text: '\t\u00a0%' },
        { text: 'BAIXAR AGORA', href: 'https://example.org/planilha', bold: true },
        { text: '%\u00a0\n' },
        { text: 'Depois', italic: true },
      ],
    }])
    expect(result.blocks).toEqual([
      { type: 'paragraph', runs: [{ text: 'Antes', bold: true }] },
      {
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'BAIXAR AGORA', url: 'https://example.org/planilha' },
      },
      { type: 'paragraph', runs: [{ text: 'Depois', italic: true }] },
    ])
    expect(result.warnings).toEqual([])
  })

  it('converts multiple marker lines while retaining inline markers and other lines unchanged', () => {
    const original = { type: 'paragraph' as const, runs: [{ text: 'Antes\n%UM%\nClique %INLINE% hoje\n%DOIS%\nDepois' }] }
    const result = resolveComponentMarkers([original])
    expect(result.blocks.map((block) => block.type === 'component' ? block.initial?.text : block.runs.map((run) => run.text).join('')))
      .toEqual(['Antes', 'UM', 'Clique %INLINE% hoje', 'DOIS', 'Depois'])
    expect(result.blocks.filter((block) => block.type === 'component').map((block) => block.instanceId))
      .toEqual(['word-button-1', 'word-button-2'])
    expect(result.warnings).toEqual([
      expect.stringContaining('%UM%'),
      expect.stringContaining('%INLINE%'),
      expect.stringContaining('%DOIS%'),
    ])
  })

  it('keeps URLs separate for multiple generic markers and uses # for a missing URL', () => {
    const result = resolveComponentMarkers([
      { type: 'paragraph', runs: [{ text: '%SAIBA MAIS%', href: 'https://example.org/mais' }] },
      paragraph('Texto entre'),
      paragraph('%CADASTRO%'),
    ])
    expect(result.blocks[0]).toMatchObject({
      componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'SAIBA MAIS', url: 'https://example.org/mais' },
    })
    expect(result.blocks[1]).toEqual(paragraph('Texto entre'))
    expect(result.blocks[2]).toMatchObject({
      componentId: 'word-button', instanceId: 'word-button-2',
      initial: { text: 'CADASTRO', url: '#' },
    })
    expect(result.warnings).toEqual([expect.stringContaining('não possui hyperlink')])
    expect(result.warnings[0]).toContain('usará #')
  })

  it('does not turn inline, empty, or multiple markers into buttons', () => {
    const inline = paragraph('Use %BAIXAR AGORA% hoje')
    const empty = paragraph('%%')
    const multiple = paragraph('%UM% %DOIS%')
    const result = resolveComponentMarkers([inline, empty, multiple])
    expect(result.blocks).toEqual([inline, empty, multiple])
    expect(result.warnings).toHaveLength(3)
    expect(result.warnings.some((warning) => warning.includes('%BAIXAR AGORA%'))).toBe(true)
  })

  it('uses # when linked runs disagree or use an unsafe protocol', () => {
    const result = resolveComponentMarkers([
      { type: 'paragraph', runs: [
        { text: '%BAIXAR', href: 'https://example.org/a' },
        { text: ' AGORA%', href: 'https://example.org/b' },
      ] },
      { type: 'paragraph', runs: [{ text: '%CADASTRO%', href: 'javascript:alert(1)' }] },
    ])
    expect(result.blocks[0]).toMatchObject({ initial: { text: 'BAIXAR AGORA', url: '#' } })
    expect(result.blocks[1]).toMatchObject({ initial: { text: 'CADASTRO', url: '#' } })
    expect(result.warnings).toHaveLength(2)
  })

  it('preserves a marker whose label cannot be saved and uses # for an oversized URL', () => {
    const overlong = paragraph(`%${'A'.repeat(1_001)}%`)
    const result = resolveComponentMarkers([
      overlong,
      { type: 'paragraph', runs: [{ text: '%ARQUIVO%', href: `https://example.org/${'x'.repeat(4_100)}` }] },
    ])
    expect(result.blocks[0]).toBe(overlong)
    expect(result.blocks[1]).toMatchObject({ initial: { text: 'ARQUIVO', url: '#' } })
    expect(result.warnings).toHaveLength(2)
  })

  it('uses # for a malformed Word hyperlink but retains a valid HTTP(S) link', () => {
    const result = resolveComponentMarkers([
      { type: 'paragraph', runs: [{ text: '%BAIXAR AGORA%', href: 'https://www/' }] },
      { type: 'paragraph', runs: [{ text: '%CADASTRO%', href: 'https://example.org/cadastro' }] },
    ])
    expect(result.blocks[0]).toMatchObject({ initial: { text: 'BAIXAR AGORA', url: '#' } })
    expect(result.blocks[1]).toMatchObject({ initial: { text: 'CADASTRO', url: 'https://example.org/cadastro' } })
    expect(result.warnings).toEqual([expect.stringContaining('usará #')])
  })

  it('is idempotent and reserves IDs of existing buttons while migrating old projects', () => {
    const existing = {
      type: 'component' as const, runs: [] as [], componentId: 'word-button', instanceId: 'word-button-2',
      initial: { text: 'EXISTENTE', url: 'https://example.org/existente' },
    }
    const result = resolveComponentMarkers([paragraph('%NOVO%'), existing, paragraph('%OUTRO%')])
    expect(result.blocks.map((block) => block.type === 'component' ? block.instanceId : block.type))
      .toEqual(['word-button-3', 'word-button-2', 'word-button-4'])
    expect(resolveComponentMarkers(result.blocks).blocks).toEqual(result.blocks)
  })
})

describe('custom component HTML', () => {
  it('extracts its defaults and renders the exact original template by default', () => {
    const definition = getComponentDefinition('acessar')
    expect(definition?.defaults).toEqual({
      text: 'Teste doc',
      url: 'https://pages.xpi.com.br/content/dam/infomoney/TESTE%20BUIDER%20MAIL/JORNADA%20DE%20PFI%20-%20Convers%C3%A3o.docx',
    })
    expect(renderCustomComponent('acessar')).toBe(ACESSAR_HTML)
  })

  it('changes only href and visible text while retaining tables, styles, and other attributes', () => {
    const url = 'https://example.org/oferta?a=1&b=2'
    const result = renderCustomComponent('acessar', { text: 'Acessar & conhecer', url })
    const expected = ACESSAR_HTML
      .replace('href="https://pages.xpi.com.br/content/dam/infomoney/TESTE%20BUIDER%20MAIL/JORNADA%20DE%20PFI%20-%20Convers%C3%A3o.docx"', 'href="https://example.org/oferta?a=1&amp;b=2"')
      .replace('>Teste doc</a>', '>Acessar &amp; conhecer</a>')
    expect(result).toBe(expected)
    expect(result).toContain('title="TEXTO DO BOTÃO" alias="TEXTO DO BOTÃO"')
    expect(result).toContain('padding:20px;border-radius:7px;')
  })

  it('uses the original table HTML for generic buttons without changing any other attributes', () => {
    const url = 'https://example.org/planilha?a=1&b=2'
    const actual = renderCustomComponent('word-button', { text: 'BAIXAR AGORA', url })
    const expected = ACESSAR_HTML
      .replace('href="https://pages.xpi.com.br/content/dam/infomoney/TESTE%20BUIDER%20MAIL/JORNADA%20DE%20PFI%20-%20Convers%C3%A3o.docx"', 'href="https://example.org/planilha?a=1&amp;b=2"')
      .replace('>Teste doc</a>', '>BAIXAR AGORA</a>')
    expect(actual).toBe(expected)
  })

  it('renders # for an empty or placeholder destination in preview and export', () => {
    const expected = ACESSAR_HTML
      .replace('href="https://pages.xpi.com.br/content/dam/infomoney/TESTE%20BUIDER%20MAIL/JORNADA%20DE%20PFI%20-%20Convers%C3%A3o.docx"', 'href="#"')
      .replace('>Teste doc</a>', '>CADASTRO</a>')
    expect(renderCustomComponentDraft('word-button', { text: 'CADASTRO', url: '' })).toBe(expected)
    expect(renderCustomComponent('word-button', { text: 'CADASTRO', url: '' })).toBe(expected)
    expect(renderCustomComponent('word-button', { text: 'CADASTRO', url: '#' })).toBe(expected)
  })

  it('escapes dangerous text and rejects non-HTTP URLs', () => {
    expect(renderCustomComponent('acessar', { text: '<img src=x onerror=alert(1)>' }))
      .toContain('&lt;img src=x onerror=alert(1)&gt;</a>')
    expect(() => renderCustomComponent('acessar', { url: 'javascript:alert(1)' })).toThrow('URL HTTP ou HTTPS')
    expect(() => renderCustomComponent('acessar', { url: 'https://example.org/\nother' })).toThrow('URL HTTP ou HTTPS')
    expect(renderCustomComponent('acessar', { url: '' })).toContain('href="#"')
    expect(() => renderCustomComponent('word-button', { text: 'BAIXAR AGORA', url: 'https://www/' })).toThrow('URL HTTP ou HTTPS')
    expect(renderCustomComponent('word-button', { text: 'IPv6', url: 'https://[2001:db8::1]/arquivo' }))
      .toContain('href="https://[2001:db8::1]/arquivo"')
  })

  it('can register another marker without changing the resolver', () => {
    const definition = registerCustomComponent({
      id: 'test_future_marker',
      name: 'Marcador futuro',
      marker: '%test_future_marker%',
      html: '<table role="presentation"><tr><td><a>Saiba mais</a></td></tr></table>',
    })
    expect(definition.defaults).toEqual({ text: 'Saiba mais', url: '' })
    expect(resolveComponentMarkers([paragraph('%test_future_marker%')]).blocks[0])
      .toMatchObject({ type: 'component', componentId: 'test_future_marker' })
    expect(renderCustomComponent('test_future_marker')).toBe('<table role="presentation"><tr><td><a href="#">Saiba mais</a></td></tr></table>')
    expect(renderCustomComponent('test_future_marker', { url: 'https://example.org/?v=$&' }))
      .toContain('<a href="https://example.org/?v=$&amp;">Saiba mais</a>')
  })
})
