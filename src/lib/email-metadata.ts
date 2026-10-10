import type { ContentBlock, InlineNode, TextBlock } from './document'

export type InitialEmailMetadata = {
  metadata: ContentBlock[]
  body: ContentBlock[]
}

// Match only explicit labels in the imported document. An email's subject may
// live outside the body; it must never be manufactured from settings or copy.
const METADATA_LABEL = /^(?:SUB|ASSUNTO|PR[ÉE][\s\u2010-\u2015-]*HEADER|LINHA\s+FINA)\s*(?::|$)/iu

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

/** Finds the first body line while retaining every original character. */
function metadataPrefixInBlock(text: string): { length: number; hasLabel: boolean } {
  let start = 0
  let hasLabel = false
  const inspectLine = (end: number): number | null => {
    const line = text.slice(start, end).trim()
    if (METADATA_LABEL.test(line)) hasLabel = true
    else if (line) return start
    return null
  }
  for (const match of text.matchAll(/\r\n|[\r\n\u2028\u2029]/gu)) {
    const bodyStart = inspectLine(match.index)
    if (bodyStart !== null) return { length: bodyStart, hasLabel }
    start = match.index + match[0].length
  }
  const bodyStart = inspectLine(text.length)
  return { length: bodyStart ?? text.length, hasLabel }
}

/**
 * Keeps the header after the leading subject/preheader lines imported from Word.
 * Paragraphs are reused unless a soft line break requires a split; in that case,
 * text and inline formatting are sliced without mutating the original runs.
 * Condition boundaries and components always remain in the body, so the header
 * can never be placed inside an audience condition.
 */
export function partitionInitialMetadata(blocks: ContentBlock[]): InitialEmailMetadata {
  const metadata: ContentBlock[] = []
  let foundLabel = false
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type !== 'paragraph' && block.type !== 'heading') break
    const text = block.runs.map((run) => run.text).join('')
    const prefix = metadataPrefixInBlock(text)
    foundLabel ||= prefix.hasLabel
    if (prefix.length === 0 && text.length > 0) break
    if (prefix.length < text.length) {
      if (!foundLabel) break
      metadata.push({ ...block, runs: sliceRuns(block.runs, 0, prefix.length) })
      const remainder: TextBlock = { ...block, runs: sliceRuns(block.runs, prefix.length, text.length) }
      return { metadata, body: [remainder, ...blocks.slice(index + 1)] }
    }
    metadata.push(block)
  }
  return foundLabel
    ? { metadata, body: blocks.slice(metadata.length) }
    : { metadata: [], body: blocks }
}
