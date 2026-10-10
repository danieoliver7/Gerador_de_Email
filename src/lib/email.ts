import type { ContentBlock, ExtractedDocument, InlineNode, TextBlock } from './document'
import { renderCustomComponent } from './custom-components'
import type { ComponentBlock, ComponentOverrides } from './custom-components'
import { validateConditionBlocks } from './conditions'
import type { ConditionBoundaryBlock } from './conditions'
import { restoreHeaderSettings } from './header'
import type { HeaderSettings } from './header'
import { partitionInitialMetadata } from './email-metadata'
import { planEmailSections, sectionComment } from './email-sections'
import type { EmailSectionName, EmailSectionPlan } from './email-sections'

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

export type RenderEmailOptions = {
  audienceSimulation?: 'member' | 'non-member'
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
type CommentSlot = { token: string; comment: string }

function uniqueSlotToken(kind: string, index: number, documentText: string): string {
  let suffix = 0
  let token: string
  do {
    token = `__ATELIE_${kind}_SLOT_${index}_${suffix++}__`
  } while (documentText.includes(token))
  return token
}

function renderComment(
  name: EmailSectionName, edge: 'start' | 'end', slots: CommentSlot[], documentText: string,
): string {
  const token = uniqueSlotToken('COMMENT', slots.length, documentText)
  slots.push({ token, comment: sectionComment(name, edge) })
  return `<mj-raw>${token}</mj-raw>`
}

function renderBlocks(
  blocks: ContentBlock[], accentColor: string, compact: boolean,
  slots: ComponentSlot[], conditionSlots: ConditionSlot[], documentText: string,
  organization?: { plan: EmailSectionPlan; comments: CommentSlot[]; options: RenderEmailOptions },
): string {
  const output: string[] = []
  const starts = new Map(organization?.plan.segments.map((segment) => [segment.start, segment]))
  const ends = new Map(organization?.plan.segments.map((segment) => [segment.end, segment]))
  let insideCondition = false

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]
    if (block.type === 'condition-start' || block.type === 'condition-end') {
      insideCondition = block.type === 'condition-start'
      if (organization?.options.audienceSimulation) continue
      const token = uniqueSlotToken('CONDITION', conditionSlots.length, documentText)
      conditionSlots.push({ token, block })
      // MJML emits mj-raw between complete table rows, a safe AJO boundary.
      output.push(`<mj-raw>${token}</mj-raw>`)
      continue
    }
    if (organization?.options.audienceSimulation === 'non-member' && insideCondition) continue
    const start = starts.get(index)
    if (start && organization) output.push(renderComment(start.name, 'start', organization.comments, documentText))
    if (block.type === 'component') {
      const token = uniqueSlotToken('COMPONENT', slots.length, documentText)
      slots.push({ token, block })
      output.push(`<mj-text padding="0 0 ${compact ? 8 : 14}px 0">${token}</mj-text>`)
    } else if (block.type !== 'list-item') {
      output.push(renderBlock(block, accentColor, compact))
    } else {
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
    const end = ends.get(index + 1)
    if (end && organization) output.push(renderComment(end.name, 'end', organization.comments, documentText))
  }

  return output.join('\n')
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

function buildMjmlWithSlots(document: ExtractedDocument, settings: EmailSettings, options: RenderEmailOptions = {}): {
  mjml: string; slots: ComponentSlot[]; conditionSlots: ConditionSlot[]; commentSlots: CommentSlot[]
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
  const commentSlots: CommentSlot[] = []
  const documentText = [
    ...document.blocks.flatMap((block) => block.runs.flatMap((run) => [run.text, run.href ?? ''])),
    settings.fontFamily,
  ].join('')
  const parts = header.enabled
    ? partitionInitialMetadata(document.blocks)
    : { metadata: [], body: document.blocks }
  const hasMetadata = parts.metadata.length > 0
  const metadata = hasMetadata
    ? renderBlocks(parts.metadata, accentColor, compact, slots, conditionSlots, documentText)
    : ''
  // The complete model determines numbering before any preview visibility filter.
  const content = renderBlocks(parts.body, accentColor, compact, slots, conditionSlots, documentText, {
    plan: planEmailSections(parts.body), comments: commentSlots, options,
  })

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
      renderComment('CABEÇALHO', 'start', commentSlots, documentText),
      `    <mj-section padding="${compact ? '0' : hasMetadata ? '0 16px' : '24px 16px 0'}" background-color="${backgroundColor}">`,
      ...renderHeader(header, surfaceColor),
      '    </mj-section>',
      renderComment('CABEÇALHO', 'end', commentSlots, documentText),
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
  return { mjml, slots, conditionSlots, commentSlots }
}

export function buildMjml(document: ExtractedDocument, settings: EmailSettings): string {
  return buildMjmlWithSlots(document, settings).mjml
}

/** Replace positions in the compiled output once, never scanning inserted button HTML. */
function fillCompiledSlots(html: string, replacements: Array<{ token: string; value: string }>): string {
  const positions = replacements.map(({ token, value }) => {
    const start = html.indexOf(token)
    if (start < 0 || html.indexOf(token, start + token.length) >= 0) {
      throw new Error('Não foi possível posicionar uma seção do email com segurança.')
    }
    return { start, end: start + token.length, value }
  }).sort((left, right) => left.start - right.start)
  const output: string[] = []
  let cursor = 0
  for (const { start, end, value } of positions) {
    output.push(html.slice(cursor, start), value)
    cursor = end
  }
  output.push(html.slice(cursor))
  return output.join('')
}

/** Inserts trusted registered components after MJML compilation so their table HTML is untouched. */
export async function renderEmail(
  document: ExtractedDocument,
  settings: EmailSettings,
  componentOverrides: Record<string, Partial<ComponentOverrides>> = {},
  options: RenderEmailOptions = {},
): Promise<RenderedEmail> {
  const validationErrors = [...new Set([...(document.errors ?? []), ...validateConditionBlocks(document.blocks)])]
  if (validationErrors.length) return { mjml: '', html: '', errors: validationErrors }
  const { mjml, slots, conditionSlots, commentSlots } = buildMjmlWithSlots(document, settings, options)

  try {
    const { default: mjml2html } = await import('mjml-browser')
    const result = await mjml2html(mjml, { validationLevel: 'soft' })
    const html = fillCompiledSlots(result.html, [
      ...slots.map(({ token, block }) => ({
        token,
        value: renderCustomComponent(block.componentId, { ...block.initial, ...componentOverrides[block.instanceId] }),
      })),
      ...conditionSlots.map(({ token, block }) => ({
        token,
        value: block.type === 'condition-start'
          ? `{%#if profile.segmentMembership.get("ups").get("${block.audience}").status = "existing"%}`
          : '{%/if%}',
      })),
      ...commentSlots.map(({ token, comment }) => ({ token, value: `\n${comment}\n` })),
    ])
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
