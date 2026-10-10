import { describe, expect, it } from 'vitest'
import type { ContentBlock, TextBlock } from './document'
import type { ComponentBlock } from './custom-components'
import { registerCustomComponent } from './custom-components'
import type { ConditionBoundaryBlock } from './conditions'
import { planEmailSections, sectionComment } from './email-sections'

function paragraph(text: string): TextBlock {
  return { type: 'paragraph', runs: [{ text }] }
}

function button(index: number): ComponentBlock {
  return { type: 'component', componentId: 'acessar', instanceId: `acessar-${index}`, runs: [] }
}

function condition(type: ConditionBoundaryBlock['type'], audience = 'BASE_ATIVO'): ConditionBoundaryBlock {
  return { type, audience, runs: [] }
}

describe('planEmailSections', () => {
  it('assigns the entire body without buttons to PARTE 1', () => {
    const blocks = [paragraph('Olá'), paragraph('Assinatura')]
    expect(planEmailSections(blocks)).toEqual({
      buttonCount: 0,
      segments: [{ name: 'PARTE 1', kind: 'part', start: 0, end: 2 }],
    })
  })

  it.each([1, 2, 4, 6])('calculates numbered parts and buttons dynamically for %i buttons', (count) => {
    const blocks: ContentBlock[] = []
    for (let index = 1; index <= count; index += 1) {
      blocks.push(paragraph(`Texto ${index}`), button(index))
    }
    blocks.push(paragraph('Assinatura'))
    const plan = planEmailSections(blocks)

    expect(plan.buttonCount).toBe(count)
    expect(plan.segments.map((segment) => segment.name)).toEqual([
      ...Array.from({ length: count }, (_, index) => [`PARTE ${index + 1}`, `BOTÃO ${index + 1}`]).flat(),
      'PARTE FINAL',
    ])
    expect(plan.segments.at(-1)).toEqual({ name: 'PARTE FINAL', kind: 'part', start: count * 2, end: count * 2 + 1 })
  })

  it('does not invent an initial part, intermediate parts or a final part between adjacent buttons', () => {
    expect(planEmailSections([button(1), button(2), button(3)]).segments).toEqual([
      { name: 'BOTÃO 1', kind: 'button', start: 0, end: 1 },
      { name: 'BOTÃO 2', kind: 'button', start: 1, end: 2 },
      { name: 'BOTÃO 3', kind: 'button', start: 2, end: 3 },
    ])
    expect(planEmailSections([])).toEqual({ buttonCount: 0, segments: [] })
  })

  it('does not mark parts made only of empty runs, spaces or line breaks around a button', () => {
    const blank: TextBlock = { type: 'paragraph', runs: [] }
    const blocks = [blank, paragraph(' \n\t'), button(1), paragraph('\r\n\u00a0 '), blank]
    expect(planEmailSections(blocks)).toEqual({
      buttonCount: 1,
      segments: [{ name: 'BOTÃO 1', kind: 'button', start: 2, end: 3 }],
    })
    expect(planEmailSections([blank, paragraph(' \n')])).toEqual({ buttonCount: 0, segments: [] })
  })

  it('includes original blank blocks when they share a segment with actual content', () => {
    const blocks = [paragraph(' '), paragraph('Introdução'), paragraph('\n'), button(1), paragraph('Fim'), paragraph(' ')]
    expect(planEmailSections(blocks).segments).toEqual([
      { name: 'PARTE 1', kind: 'part', start: 0, end: 3 },
      { name: 'BOTÃO 1', kind: 'button', start: 3, end: 4 },
      { name: 'PARTE FINAL', kind: 'part', start: 4, end: 6 },
    ])
  })

  it('counts a conditional button in the first position', () => {
    const plan = planEmailSections([
      condition('condition-start'), button(1), condition('condition-end'), paragraph('Texto'), button(2),
    ])
    expect(plan.segments.map((segment) => segment.name)).toEqual(['BOTÃO 1', 'PARTE 2', 'BOTÃO 2'])
    expect(plan.segments[0]).toEqual({ name: 'BOTÃO 1', kind: 'button', start: 1, end: 2 })
    expect(plan.buttonCount).toBe(2)
  })

  it('counts conditional buttons between common buttons and across different audiences', () => {
    const blocks = [
      paragraph('Abertura'), button(1),
      condition('condition-start'), button(2), condition('condition-end'),
      condition('condition-start', 'CLIENTES_VIP'), button(3), condition('condition-end', 'CLIENTES_VIP'),
      button(4), paragraph('Assinatura'),
    ]
    const plan = planEmailSections(blocks)
    expect(plan.buttonCount).toBe(4)
    expect(plan.segments.filter((segment) => segment.kind === 'button').map((segment) => segment.name)).toEqual([
      'BOTÃO 1', 'BOTÃO 2', 'BOTÃO 3', 'BOTÃO 4',
    ])
    expect(plan.segments.at(-1)?.name).toBe('PARTE FINAL')
  })

  it('separately counts multiple buttons within one IF without extending the IF scope', () => {
    const plan = planEmailSections([
      condition('condition-start'), paragraph('Oferta'), button(1), paragraph('Mais oportunidades'), button(2),
      paragraph('Condição continua'), condition('condition-end'), paragraph('Assinatura'),
    ])
    expect(plan.segments).toEqual([
      { name: 'PARTE 1', kind: 'part', start: 1, end: 2 },
      { name: 'BOTÃO 1', kind: 'button', start: 2, end: 3 },
      { name: 'PARTE 2', kind: 'part', start: 3, end: 4 },
      { name: 'BOTÃO 2', kind: 'button', start: 4, end: 5 },
      { name: 'PARTE FINAL', kind: 'part', start: 5, end: 6 },
      { name: 'PARTE FINAL', kind: 'part', start: 7, end: 8 },
    ])
  })

  it('segments a logical part at each IF boundary, retaining its number', () => {
    const blocks = [
      paragraph('Comum antes'), condition('condition-start'), paragraph('Condicionado'), condition('condition-end'),
      paragraph('Comum depois'), button(1), paragraph('Final'),
    ]
    expect(planEmailSections(blocks).segments).toEqual([
      { name: 'PARTE 1', kind: 'part', start: 0, end: 1 },
      { name: 'PARTE 1', kind: 'part', start: 2, end: 3 },
      { name: 'PARTE 1', kind: 'part', start: 4, end: 5 },
      { name: 'BOTÃO 1', kind: 'button', start: 5, end: 6 },
      { name: 'PARTE FINAL', kind: 'part', start: 6, end: 7 },
    ])
  })

  it('segments PARTE 1 safely in an email with IFs and no buttons', () => {
    const blocks = [paragraph('Introdução'), condition('condition-start'), paragraph('Oferta'), condition('condition-end')]
    expect(planEmailSections(blocks).segments).toEqual([
      { name: 'PARTE 1', kind: 'part', start: 0, end: 1 },
      { name: 'PARTE 1', kind: 'part', start: 2, end: 3 },
    ])
  })

  it('does not classify links in paragraphs, headings or lists as buttons', () => {
    const blocks: ContentBlock[] = [
      { type: 'paragraph', runs: [{ text: 'Link', href: 'https://example.com' }] },
      { type: 'heading', runs: [{ text: 'Título', href: 'https://example.com' }] },
      { type: 'list-item', runs: [{ text: 'Item', href: 'https://example.com' }] },
      button(1),
    ]
    expect(planEmailSections(blocks)).toEqual({
      buttonCount: 1,
      segments: [
        { name: 'PARTE 1', kind: 'part', start: 0, end: 3 },
        { name: 'BOTÃO 1', kind: 'button', start: 3, end: 4 },
      ],
    })
  })

  it('counts editor button components by their registered kind and keeps content components in a part', () => {
    registerCustomComponent({
      id: 'section-test-content', name: 'Conteúdo de teste', marker: null, kind: 'content',
      html: '<table role="presentation"><tr><td><a href="https://example.com">Saiba mais</a></td></tr></table>',
    })
    registerCustomComponent({
      id: 'section-test-editor-button', name: 'Botão do editor de teste', marker: null, kind: 'button',
      html: '<table role="presentation"><tr><td><a href="https://example.com">Acessar</a></td></tr></table>',
    })
    const blocks: ContentBlock[] = [
      { type: 'component', componentId: 'section-test-content', instanceId: 'content-1', runs: [] },
      { type: 'component', componentId: 'section-test-editor-button', instanceId: 'button-1', runs: [] },
    ]
    expect(planEmailSections(blocks)).toEqual({
      buttonCount: 1,
      segments: [
        { name: 'PARTE 1', kind: 'part', start: 0, end: 1 },
        { name: 'BOTÃO 1', kind: 'button', start: 1, end: 2 },
      ],
    })
  })

  it('never mutates source blocks and produces the same plan on regeneration', () => {
    const blocks = [paragraph('Abertura'), condition('condition-start'), button(1), condition('condition-end'), paragraph('Fim')]
    const original = structuredClone(blocks)
    const first = planEmailSections(blocks)
    expect(planEmailSections(blocks)).toEqual(first)
    expect(blocks).toEqual(original)
  })

  it('recalculates all numbers after adding or removing a component', () => {
    const blocks = [paragraph('Abertura'), button(8), paragraph('Oferta'), button(20), paragraph('Fim')]
    expect(planEmailSections(blocks).segments.map((segment) => segment.name)).toEqual([
      'PARTE 1', 'BOTÃO 1', 'PARTE 2', 'BOTÃO 2', 'PARTE FINAL',
    ])
    const changed = [button(30), ...blocks.filter((block) => block !== blocks[1])]
    expect(planEmailSections(changed).segments.map((segment) => segment.name)).toEqual([
      'BOTÃO 1', 'PARTE 2', 'BOTÃO 2', 'PARTE FINAL',
    ])
  })

  it('covers every content block once and places no IF boundary inside a segment', () => {
    const blocks = [
      paragraph('Abertura'), condition('condition-start'), paragraph('Oferta'), button(1),
      paragraph('Lista'), button(2), condition('condition-end'), paragraph('Fim'), button(3),
    ]
    const covered = planEmailSections(blocks).segments.flatMap((segment) => {
      expect(segment.end).toBeGreaterThan(segment.start)
      const indices = Array.from({ length: segment.end - segment.start }, (_, offset) => segment.start + offset)
      for (const index of indices) expect(blocks[index].type).not.toMatch(/^condition-/)
      return indices
    })
    expect(covered).toEqual([0, 2, 3, 4, 5, 7, 8])
    expect(new Set(covered).size).toBe(covered.length)
  })
})

describe('sectionComment', () => {
  it('generates paired literal HTML comments for the header and numbered body sections', () => {
    expect(sectionComment('CABEÇALHO', 'start')).toBe('<!-- CABEÇALHO - INÍCIO -->')
    expect(sectionComment('CABEÇALHO', 'end')).toBe('<!-- CABEÇALHO - FIM -->')
    expect(sectionComment('PARTE 2', 'start')).toBe('<!-- PARTE 2 - INÍCIO -->')
    expect(sectionComment('BOTÃO 4', 'end')).toBe('<!-- BOTÃO 4 - FIM -->')
    expect(sectionComment('PARTE FINAL', 'end')).toBe('<!-- PARTE FINAL - FIM -->')
  })
})
