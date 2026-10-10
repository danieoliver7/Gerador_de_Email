import { describe, expect, it } from 'vitest'
import { partitionInitialMetadata } from './email-metadata'
import type { ContentBlock, TextBlock } from './document'

const paragraph = (text: string): TextBlock => ({ type: 'paragraph', runs: [{ text }] })
const textOf = (blocks: ContentBlock[]) => blocks.map((block) => block.runs.map((run) => run.text).join(''))

describe('partitionInitialMetadata', () => {
  it.each(['Pré-header', 'Pre-header', 'Preheader', 'préheader', 'Pré header', 'Pré‑header', 'Linha fina'])(
    'keeps the initial subject and %s before the body without rewriting the imported paragraphs',
    (label) => {
      const blocks = [paragraph('Assunto: Oferta'), paragraph(`${label}: Confira os detalhes`), paragraph('Olá, cliente!')]
      const result = partitionInitialMetadata(blocks)
      expect(result.metadata).toEqual(blocks.slice(0, 2))
      expect(result.metadata[0]).toBe(blocks[0])
      expect(result.metadata[1]).toBe(blocks[1])
      expect(result.body).toEqual([blocks[2]])
      expect(result.body[0]).toBe(blocks[2])
    },
  )

  it('supports SUB, case differences, blank paragraphs and metadata represented as headings', () => {
    const subject: TextBlock = { type: 'heading', level: 3, runs: [{ text: 'sUb : ' }, { text: 'Oferta', bold: true }] }
    const blocks = [paragraph('  '), subject, paragraph(''), paragraph('pReHeAdEr: Detalhes'), paragraph('Corpo')]
    const result = partitionInitialMetadata(blocks)
    expect(result.metadata).toEqual(blocks.slice(0, 4))
    expect(result.body).toEqual([blocks[4]])
    expect(result.metadata[1]).toBe(subject)
  })

  it('places only a leading preheader in metadata when the subject exists outside the Word body', () => {
    const result = partitionInitialMetadata([paragraph('Pré-header: Detalhes'), paragraph('Conteúdo')])
    expect(textOf(result.metadata)).toEqual(['Pré-header: Detalhes'])
    expect(textOf(result.body)).toEqual(['Conteúdo'])
    expect(textOf([...result.metadata, ...result.body]).join('')).not.toContain('Assunto')
  })

  it('keeps ordinary prose and subsequent metadata-like labels in their original body position', () => {
    const blocks = [paragraph('Assunto da campanha'), paragraph('Pré-header: Isto pertence ao corpo')]
    const result = partitionInitialMetadata(blocks)
    expect(result.metadata).toEqual([])
    expect(result.body).toBe(blocks)

    const prefixed = [paragraph('Assunto: Oferta'), paragraph('Corpo'), paragraph('Preheader: Outro texto')]
    expect(partitionInitialMetadata(prefixed)).toEqual({ metadata: [prefixed[0]], body: prefixed.slice(1) })
  })

  it('returns the untouched body when no explicit metadata label is present', () => {
    const blocks = [paragraph(''), paragraph('Confira nossas ofertas'), paragraph('Mensagem final')]
    expect(partitionInitialMetadata(blocks).metadata).toEqual([])
    expect(partitionInitialMetadata(blocks).body).toBe(blocks)
    expect(partitionInitialMetadata([])).toEqual({ metadata: [], body: [] })
  })

  it.each(['\n', '\r\n', '\r', '\u2028', '\u2029'])(
    'splits metadata and copy separated by %j while retaining every character and formatting mark',
    (separator) => {
      const block: TextBlock = {
        type: 'paragraph',
        runs: [
          { text: 'Assunto: ' },
          { text: `Oferta${separator}Pré-header: `, bold: true },
          { text: `Detalhes${separator}Olá `, italic: true },
          { text: 'cliente', href: 'https://example.org/oferta' },
        ],
      }
      const original = structuredClone(block)
      const trailing = paragraph('Final')
      const result = partitionInitialMetadata([block, trailing])
      expect(textOf(result.metadata)).toEqual([`Assunto: Oferta${separator}Pré-header: Detalhes${separator}`])
      expect(textOf(result.body)).toEqual(['Olá cliente', 'Final'])
      expect([...result.metadata[0].runs, ...result.body[0].runs].map((run) => run.text).join(''))
        .toBe(block.runs.map((run) => run.text).join(''))
      expect(result.metadata[0].runs).toContainEqual({ text: `Oferta${separator}Pré-header: `, bold: true })
      expect(result.metadata[0].runs).toContainEqual({ text: `Detalhes${separator}`, italic: true })
      expect(result.body[0].runs).toEqual([
        { text: 'Olá ', italic: true }, { text: 'cliente', href: 'https://example.org/oferta' },
      ])
      expect(result.body[1]).toBe(trailing)
      expect(block).toEqual(original)
    },
  )

  it('handles labels and a CRLF separator spanning separate inline runs', () => {
    const block: TextBlock = {
      type: 'heading', level: 2,
      runs: [{ text: 'Assunto: Oferta\r' }, { text: '\nPré-', bold: true }, { text: 'header: Detalhes\r\nCorpo', italic: true }],
    }
    const result = partitionInitialMetadata([block])
    expect(textOf(result.metadata)).toEqual(['Assunto: Oferta\r\nPré-header: Detalhes\r\n'])
    expect(result.metadata[0]).toMatchObject({ type: 'heading', level: 2 })
    expect(result.body[0]).toEqual({ type: 'heading', level: 2, runs: [{ text: 'Corpo', italic: true }] })
  })

  it.each<ContentBlock>([
    { type: 'condition-start', runs: [], audience: 'BASE_ATIVO' },
    { type: 'condition-end', runs: [], audience: 'BASE_ATIVO' },
    { type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' },
    { type: 'list-item', runs: [{ text: 'Assunto: Item de lista' }], ordered: false },
  ])('never crosses a $type boundary while locating the header position', (boundary) => {
    const blocks = [paragraph('Assunto: Oferta'), boundary, paragraph('Pré-header: Texto após limite')]
    expect(partitionInitialMetadata(blocks)).toEqual({ metadata: [blocks[0]], body: blocks.slice(1) })

    const initial = [boundary, paragraph('Assunto: Dentro do corpo')]
    expect(partitionInitialMetadata(initial).metadata).toEqual([])
    expect(partitionInitialMetadata(initial).body).toBe(initial)
  })

  it('keeps an existing separator paragraph in its original order without generating new content', () => {
    const blocks = [paragraph('Assunto: Oferta'), paragraph('Pré-header: Detalhes'), paragraph('________________'), paragraph('Corpo')]
    const result = partitionInitialMetadata(blocks)
    expect(result.metadata).toEqual(blocks.slice(0, 2))
    expect(result.body).toEqual(blocks.slice(2))
    expect([...result.metadata, ...result.body]).toEqual(blocks)
  })
})
