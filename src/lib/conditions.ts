import type { ContentBlock, InlineNode, TextBlock } from './document'

/** A condition stays in the block model until MJML has produced complete email rows. */
export type ConditionBoundaryBlock =
  | { type: 'condition-start'; runs: []; audience: string }
  | { type: 'condition-end'; runs: []; audience: string }

export type ResolvedConditions = {
  blocks: ContentBlock[]
  errors: string[]
}

const IDENTIFIER = /^[A-Za-z0-9_]+$/
const RAW_MARKER = /##[^\r\n\u2028\u2029]*?##/gu

export function validAudience(value: string): boolean {
  return value.length > 0 && value.length <= 128 && IDENTIFIER.test(value)
}

type Line = { start: number; end: number; after: number; before: number }

function linesOf(text: string): Line[] {
  const lines: Line[] = []
  let start = 0
  let previousBreakStart = 0
  for (const match of text.matchAll(/\r\n|[\r\n\u2028\u2029]/gu)) {
    const end = match.index
    const after = end + match[0].length
    lines.push({ start, end, after, before: start === 0 ? 0 : previousBreakStart })
    previousBreakStart = end
    start = after
  }
  lines.push({ start, end: text.length, after: text.length, before: start === 0 ? 0 : previousBreakStart })
  return lines
}

function sliceRuns(runs: InlineNode[], start: number, end: number): InlineNode[] {
  const sliced: InlineNode[] = []
  let offset = 0
  for (const run of runs) {
    const next = offset + run.text.length
    if (next > start && offset < end) {
      const text = run.text.slice(Math.max(0, start - offset), Math.min(run.text.length, end - offset))
      if (text) sliced.push({ ...run, text })
    }
    offset = next
  }
  return sliced
}

function parseMarker(value: string): ConditionBoundaryBlock | null {
  const match = /^##(.*)##$/u.exec(value)
  if (!match) return null
  const closing = match[1].startsWith('FIM_')
  const audience = closing ? match[1].slice(4) : match[1]
  if (!validAudience(audience)) return null
  return { type: closing ? 'condition-end' : 'condition-start', runs: [], audience }
}

function markerErrors(text: string, context: string): string[] {
  const markers = [...text.matchAll(RAW_MARKER)]
  if (markers.length) return markers.map((match) => `${match[0]}: ${context}`)
  return text.includes('##') ? [`${text.trim().slice(0, 100)}: marcador de audiência incompleto ou inválido.`] : []
}

/** Checks pairing and placement without changing any content or resolving markers. */
export function validateConditionBlocks(blocks: ContentBlock[]): string[] {
  const errors: string[] = []
  let open: { audience: string; position: number; hasContent: boolean } | null = null
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'condition-start' || block.type === 'condition-end') {
      if (!validAudience(block.audience) || block.runs.length !== 0) {
        errors.push(`Bloco condicional ${index + 1}: identificador de audiência inválido.`)
        continue
      }
      if (block.type === 'condition-start') {
        if (open) errors.push(`##${block.audience}## está dentro de ##${open.audience}##. Condições aninhadas ou aberturas duplicadas não são permitidas.`)
        else open = { audience: block.audience, position: index, hasContent: false }
      } else if (!open) {
        errors.push(`##FIM_${block.audience}## não possui marcador de abertura.`)
      } else if (open.audience !== block.audience) {
        errors.push(`##FIM_${block.audience}## não corresponde a ##${open.audience}##.`)
      } else {
        if (!open.hasContent) errors.push(`##${open.audience}## e ##FIM_${open.audience}## delimitam um bloco vazio.`)
        open = null
      }
      continue
    }
    if (open) open.hasContent = true
    const text = block.runs.map((run) => run.text).join('')
    if (text.includes('##')) errors.push(...markerErrors(text, 'deixe o marcador sozinho em um parágrafo ou linha do Word; o texto foi preservado.'))
  }
  if (open) errors.push(`##${open.audience}## não possui ##FIM_${open.audience}## correspondente.`)
  return [...new Set(errors)]
}

/** Resolve standalone Word markers, including characters split across DOCX runs. */
export function resolveConditionMarkers(blocks: ContentBlock[]): ResolvedConditions {
  const resolved: ContentBlock[] = []
  const errors: string[] = []

  for (const block of blocks) {
    if (block.type !== 'paragraph') {
      resolved.push(block)
      continue
    }
    const text = block.runs.map((run) => run.text).join('')
    const matches: Array<{ line: Line; boundary: ConditionBoundaryBlock }> = []
    for (const line of linesOf(text)) {
      const value = text.slice(line.start, line.end).trim()
      if (!value.includes('##')) continue
      const boundary = parseMarker(value)
      if (boundary) matches.push({ line, boundary })
      // Invalid or inline markers remain in the original text; validation reports them.
    }
    if (matches.length === 0) {
      resolved.push(block)
      continue
    }

    let cursor = 0
    for (const { line, boundary } of matches) {
      const end = line.start === 0 ? 0 : line.before
      if (end > cursor) resolved.push({ ...block, runs: sliceRuns(block.runs, cursor, end) } as TextBlock)
      resolved.push(boundary)
      cursor = line.after
    }
    if (cursor < text.length) resolved.push({ ...block, runs: sliceRuns(block.runs, cursor, text.length) } as TextBlock)
  }

  errors.push(...validateConditionBlocks(resolved))
  return { blocks: resolved, errors: [...new Set(errors)] }
}
