import type { ContentBlock, ExtractedDocument, InlineNode, TextBlock } from './document'
import { renderCustomComponent } from './custom-components'
import type { ComponentBlock, ComponentOverrides } from './custom-components'
import { validateConditionBlocks } from './conditions'
import type { ConditionBoundaryBlock } from './conditions'
import { restoreHeaderSettings } from './header'
import type { HeaderSettings } from './header'

export type EmailSettings = {
  width: number
  fontFamily: string
  backgroundColor: string
  surfaceColor: string
  textColor: string
  accentColor: string
  template: 'classic' | 'compact'
  header: HeaderSettings
}

export type RenderedEmail = {
  mjml: string
  html: string
  errors: string[]
}

const DEFAULT_COLOR = {
  backgroundColor: '#f3f4f6',
  surfaceColor: '#ffffff',
  textColor: '#1f2937',
  accentColor: '#2563eb',
} as const

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function safeColor(value: string, fallback: string): string {
  return /^#[0-9a-fA-F]{6}$/.test(value) ? value : fallback
}

function safeFontFamily(value: string): string {
  const trimmed = value.trim()
  return trimmed.length > 0 && trimmed.length <= 120 && /^[\w\s,'"-]+$/.test(trimmed)
    ? trimmed
    : 'Arial, Helvetica, sans-serif'
}

function safeHref(value: string | undefined): string | undefined {
  if (!value) return undefined
  const href = value.trim()
  if (/\s|[\u0000-\u001f\u007f]/.test(href)) return undefined

  if (/^#[A-Za-z0-9._-]+$/.test(href)) return href

  if (/^https?:\/\//i.test(href)) {
    try {
      const parsed = new URL(href)
      return parsed.hostname ? href : undefined
    } catch {
      return undefined
    }
  }

  if (/^(mailto|tel):[^\s]+$/i.test(href)) return href
  return undefined
}

function renderRun(run: InlineNode, accentColor: string): string {
  let content = escapeHtml(run.text).replace(/\r\n?|\n/g, '<br />')
  if (run.bold) content = `<strong>${content}</strong>`
  if (run.italic) content = `<em>${content}</em>`

  const href = safeHref(run.href)
  if (href) {
    content = `<a href="${escapeHtml(href)}" style="color:${accentColor};text-decoration:underline;">${content}</a>`
  }

  return content
}

function renderRuns(runs: InlineNode[], accentColor: string): string {
  return runs.map((run) => renderRun(run, accentColor)).join('') || '&nbsp;'
}

function renderList(blocks: TextBlock[], accentColor: string, compact: boolean): string {
  const ordered = Boolean(blocks[0].ordered)
  const level = Math.min(5, Math.max(0, blocks[0].level ?? 0))
  const tag = ordered ? 'ol' : 'ul'
  const items = blocks
    .map((block) => `<li style="margin:0 0 ${compact ? 4 : 8}px 0;">${renderRuns(block.runs, accentColor)}</li>`)
    .join('')

  return `<mj-text padding="0 0 ${compact ? 8 : 14}px 0"><${tag} style="margin:0;padding-left:${20 + level * 16}px;">${items}</${tag}></mj-text>`
}

function renderBlock(block: TextBlock, accentColor: string, compact: boolean): string {
  const content = renderRuns(block.runs, accentColor)
  if (block.type === 'heading') {
    const level = Math.min(6, Math.max(1, block.level ?? 1))
    const fontSize = Math.max(18, 30 - (level - 1) * 2)
    return `<mj-text font-size="${fontSize}px" font-weight="700" line-height="1.25" padding="0 0 ${compact ? 10 : 18}px 0">${content}</mj-text>`
  }

  return `<mj-text padding="0 0 ${compact ? 8 : 14}px 0">${content}</mj-text>`
}

type ComponentSlot = { token: string; block: ComponentBlock }
type ConditionSlot = { token: string; block: ConditionBoundaryBlock }

function renderBlocks(
  blocks: ContentBlock[], accentColor: string, compact: boolean,
  slots: ComponentSlot[], conditionSlots: ConditionSlot[], documentText: string,
): string {
  const output: string[] = []

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'condition-start' || block.type === 'condition-end') {
      let suffix = 0
      let token = ''
      do {
        token = `__ATELIE_CONDITION_SLOT_${index}_${suffix}__`
        suffix += 1
      } while (documentText.includes(token))
      conditionSlots.push({ token, block })
      // MJML emits mj-raw between complete table rows, a safe AJO boundary.
      output.push(`<mj-raw>${token}</mj-raw>`)
      continue
    }
    if (block.type === 'component') {
      let suffix = 0
      let token = ''
      do {
        token = `__ATELIE_COMPONENT_SLOT_${index}_${suffix}__`
        suffix += 1
      } while (documentText.includes(token))
      slots.push({ token, block })
      output.push(`<mj-text padding="0 0 ${compact ? 8 : 14}px 0">${token}</mj-text>`)
      continue
    }
    if (block.type !== 'list-item') {
      output.push(renderBlock(block, accentColor, compact))
      continue
    }

    const group = [block]
    while (
      index + 1 < blocks.length &&
      blocks[index + 1].type === 'list-item' &&
      Boolean((blocks[index + 1] as TextBlock).ordered) === Boolean(block.ordered) &&
      ((blocks[index + 1] as TextBlock).level ?? 0) === (block.level ?? 0)
    ) {
      group.push(blocks[++index] as TextBlock)
    }
    output.push(renderList(group, accentColor, compact))
  }

  return output.join('\n')
}

const METADATA_LABEL = /^(?:(?:SUB|ASSUNTO)\s*(?::|$)|LINHA\s+FINA\s*(?::|$))/iu

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

/** Finds the first body line without changing any Word text or inline marks. */
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

/** Only a leading labelled Word prefix precedes the header, including soft line breaks. */
function partitionInitialMetadata(blocks: ContentBlock[]): { metadata: ContentBlock[]; body: ContentBlock[] } {
  const metadata: ContentBlock[] = []
  let foundLabel = false
  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type !== 'paragraph' && block.type !== 'heading') break
    const text = block.runs.map((run) => run.text).join('')
    const prefix = metadataPrefixInBlock(text)
    foundLabel ||= prefix.hasLabel
    if (prefix.length === 0) break
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

/** A zero-padded MJML table keeps the empty header at the content area's full width. */
function renderHeader(header: HeaderSettings, surfaceColor: string): string[] {
  const color = safeColor(header.backgroundColor, '#000000')
  const height = header.height
  return [
    `      <mj-column background-color="${surfaceColor}" padding="0">`,
    ...(header.marginTop > 0 ? [`        <mj-spacer height="${header.marginTop}px" padding="0" />`] : []),
    '        <mj-table padding="0" width="100%" border="0" cellpadding="0" cellspacing="0">',
    `          <tr><td align="center" bgcolor="${color}" height="${height}" style="background-color:${color};height:${height}px;font-size:0;line-height:0;mso-line-height-rule:exactly;">&nbsp;</td></tr>`,
    '        </mj-table>',
    ...(header.marginBottom > 0 ? [`        <mj-spacer height="${header.marginBottom}px" padding="0" />`] : []),
    '      </mj-column>',
  ]
}

function buildMjmlWithSlots(document: ExtractedDocument, settings: EmailSettings): {
  mjml: string; slots: ComponentSlot[]; conditionSlots: ConditionSlot[]
} {
  const width = Number.isFinite(settings.width)
    ? Math.min(800, Math.max(320, Math.round(settings.width)))
    : 600
  const fontFamily = escapeHtml(safeFontFamily(settings.fontFamily))
  const backgroundColor = safeColor(settings.backgroundColor, DEFAULT_COLOR.backgroundColor)
  const surfaceColor = safeColor(settings.surfaceColor, DEFAULT_COLOR.surfaceColor)
  const textColor = safeColor(settings.textColor, DEFAULT_COLOR.textColor)
  const accentColor = safeColor(settings.accentColor, DEFAULT_COLOR.accentColor)
  const compact = settings.template === 'compact'
  const header = restoreHeaderSettings(settings.header)
  const slots: ComponentSlot[] = []
  const conditionSlots: ConditionSlot[] = []
  const documentText = document.blocks.flatMap((block) => block.runs.map((run) => run.text)).join('')
  const parts = header.enabled
    ? partitionInitialMetadata(document.blocks)
    : { metadata: [], body: document.blocks }
  const hasMetadata = parts.metadata.length > 0
  const metadata = hasMetadata
    ? renderBlocks(parts.metadata, accentColor, compact, slots, conditionSlots, documentText)
    : ''
  const content = renderBlocks(parts.body, accentColor, compact, slots, conditionSlots, documentText)

  // With the header enabled, adjacent sections share the same outer inset. The
  // header column and table have zero padding, so its color reaches both edges.
  const body = header.enabled
    ? [
      ...(hasMetadata ? [
        `    <mj-section padding="${compact ? '0' : '24px 16px 0'}" background-color="${backgroundColor}">`,
        `      <mj-column background-color="${surfaceColor}" padding="${compact ? '14px 18px 0' : '28px 26px 0'}">`,
        metadata,
        '      </mj-column>',
        '    </mj-section>',
      ] : []),
      `    <mj-section padding="${compact ? '0' : hasMetadata ? '0 16px' : '24px 16px 0'}" background-color="${backgroundColor}">`,
      ...renderHeader(header, surfaceColor),
      '    </mj-section>',
      `    <mj-section padding="${compact ? '0' : '0 16px 24px'}" background-color="${backgroundColor}">`,
      `      <mj-column background-color="${surfaceColor}" padding="${compact ? '14px 18px' : '28px 26px'}">`,
      content,
      '      </mj-column>',
      '    </mj-section>',
    ]
    : [
      `    <mj-section padding="${compact ? '0' : '24px 16px'}" background-color="${backgroundColor}">`,
      `      <mj-column background-color="${surfaceColor}" padding="${compact ? '14px 18px' : '28px 26px'}">`,
      content,
      '      </mj-column>',
      '    </mj-section>',
    ]

  const mjml = [
    '<mjml>',
    '  <mj-head>',
    `    <mj-attributes><mj-text font-family="${fontFamily}" font-size="16px" color="${textColor}" line-height="1.5" /></mj-attributes>`,
    '  </mj-head>',
    `  <mj-body width="${width}px" background-color="${backgroundColor}">`,
    ...body,
    '  </mj-body>',
    '</mjml>',
  ].join('\n')
  return { mjml, slots, conditionSlots }
}

export function buildMjml(document: ExtractedDocument, settings: EmailSettings): string {
  return buildMjmlWithSlots(document, settings).mjml
}

/** Inserts trusted registered components after MJML compilation so their table HTML is untouched. */
export async function renderEmail(
  document: ExtractedDocument,
  settings: EmailSettings,
  componentOverrides: Record<string, Partial<ComponentOverrides>> = {},
): Promise<RenderedEmail> {
  const validationErrors = [...new Set([...(document.errors ?? []), ...validateConditionBlocks(document.blocks)])]
  if (validationErrors.length) return { mjml: '', html: '', errors: validationErrors }
  const { mjml, slots, conditionSlots } = buildMjmlWithSlots(document, settings)

  try {
    const { default: mjml2html } = await import('mjml-browser')
    const result = await mjml2html(mjml, { validationLevel: 'soft' })
    let html = result.html
    for (const { token, block } of slots) {
      if (!html.includes(token)) throw new Error(`Não foi possível posicionar o componente ${block.componentId} no HTML.`)
      html = html.replace(token, renderCustomComponent(block.componentId, {
        ...block.initial,
        ...componentOverrides[block.instanceId],
      }))
    }
    for (const { token, block } of conditionSlots) {
      if (!html.includes(token)) throw new Error(`Não foi possível posicionar a condição ${block.audience} no HTML.`)
      const expression = block.type === 'condition-start'
        ? `{%#if profile.segmentMembership.get("ups").get("${block.audience}").status = "existing"%}`
        : '{%/if%}'
      html = html.replace(token, expression)
    }
    return {
      mjml,
      html,
      errors: result.errors.map((error) => error.formattedMessage || error.message),
    }
  } catch (error) {
    return {
      mjml,
      html: '',
      errors: [error instanceof Error ? error.message : String(error)],
    }
  }
}
