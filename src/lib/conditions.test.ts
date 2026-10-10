import { describe, expect, it } from 'vitest'
import { parseMammothHtml } from './document'
import { resolveConditionMarkers, validateConditionBlocks } from './conditions'
import type { ContentBlock } from './document'

function paragraph(text: string): ContentBlock {
  return { type: 'paragraph', runs: [{ text }] }
}

describe('audience condition markers', () => {
  it('resolves one condition from fragmented runs while preserving formatted copy and outside content', () => {
    const parsed = parseMammothHtml('<p>Antes</p><p><strong>##BASE</strong>_ATIVO##</p><p>Texto <strong>forte</strong> e <a href="https://example.org">link</a></p><p>##FIM_BASE_ATIVO##</p><p>Depois</p>')
    expect(parsed.errors).toEqual([])
    expect(parsed.blocks.map((block) => block.type)).toEqual([
      'paragraph', 'condition-start', 'paragraph', 'condition-end', 'paragraph',
    ])
    expect(parsed.blocks[1]).toEqual({ type: 'condition-start', runs: [], audience: 'BASE_ATIVO' })
    expect(parsed.blocks[2].runs).toContainEqual({ text: 'forte', bold: true })
    expect(parsed.blocks[2].runs).toContainEqual({ text: 'link', href: 'https://example.org' })
    expect(parsed.blocks[3]).toEqual({ type: 'condition-end', runs: [], audience: 'BASE_ATIVO' })
  })

  it('keeps different and repeated audiences as independent boundaries', () => {
    const parsed = parseMammothHtml('<p>##BASE_ATIVO##</p><p>Um</p><p>##FIM_BASE_ATIVO##</p><p>Entre</p><p>##CLIENTES_VIP##</p><p>Dois</p><p>##FIM_CLIENTES_VIP##</p><p>##BASE_ATIVO##</p><p>Três</p><p>##FIM_BASE_ATIVO##</p>')
    expect(parsed.errors).toEqual([])
    expect(parsed.blocks.filter((block) => block.type === 'condition-start').map((block) => block.audience))
      .toEqual(['BASE_ATIVO', 'CLIENTES_VIP', 'BASE_ATIVO'])
    expect(parsed.blocks[3]).toEqual(paragraph('Entre'))
  })

  it('keeps a custom button between the exact condition boundaries', () => {
    const parsed = parseMammothHtml('<p>##CLIENTES_VIP##</p><p>Oferta</p><p>%acessar%</p><p>##FIM_CLIENTES_VIP##</p>')
    expect(parsed.errors).toEqual([])
    expect(parsed.blocks.map((block) => block.type)).toEqual([
      'condition-start', 'paragraph', 'component', 'condition-end',
    ])
  })

  it('accepts a marker on its own line within a Word paragraph', () => {
    const parsed = parseMammothHtml('<p>Antes<br><strong>##CLI</strong>ENTES_VIP##<br>Dentro<br>##FIM_CLIENTES_VIP##<br>Depois</p>')
    expect(parsed.errors).toEqual([])
    expect(parsed.blocks.map((block) => block.type)).toEqual([
      'paragraph', 'condition-start', 'paragraph', 'condition-end', 'paragraph',
    ])
    expect(parsed.blocks[0].runs.map((run) => run.text).join('')).toBe('Antes')
    expect(parsed.blocks[2].runs.map((run) => run.text).join('')).toBe('Dentro')
    expect(parsed.blocks[4].runs.map((run) => run.text).join('')).toBe('Depois')
  })

  it('rejects mismatched, unmatched, nested, empty and inline markers without inventing an audience', () => {
    const mismatch = parseMammothHtml('<p>##BASE_ATIVO##</p><p>Texto</p><p>##FIM_CLIENTES_VIP##</p>')
    expect(mismatch.errors?.join(' ')).toMatch(/não corresponde/)
    const closing = parseMammothHtml('<p>##FIM_BASE_ATIVO##</p>')
    expect(closing.errors?.join(' ')).toMatch(/não possui marcador de abertura/)
    const nested = parseMammothHtml('<p>##BASE_ATIVO##</p><p>##CLIENTES_VIP##</p><p>Texto</p><p>##FIM_CLIENTES_VIP##</p><p>##FIM_BASE_ATIVO##</p>')
    expect(nested.errors?.join(' ')).toMatch(/aninhadas/)
    const empty = parseMammothHtml('<p>##BASE_ATIVO##</p><p>##FIM_BASE_ATIVO##</p>')
    expect(empty.errors?.join(' ')).toMatch(/vazio/)
    const invalid = parseMammothHtml('<p>##BASE-ATIVO##</p><p>##BASE_ATIVO## no meio</p>')
    expect(invalid.blocks).toEqual([paragraph('##BASE-ATIVO##'), paragraph('##BASE_ATIVO## no meio')])
    expect(invalid.errors?.join(' ')).toMatch(/##BASE-ATIVO##/)
    expect(invalid.errors?.join(' ')).toMatch(/##BASE_ATIVO##/)
    expect(validateConditionBlocks(resolveConditionMarkers([paragraph('##CLIENTES_VIP##')]).blocks).join(' '))
      .toMatch(/não possui ##FIM_CLIENTES_VIP##/)
  })
})
