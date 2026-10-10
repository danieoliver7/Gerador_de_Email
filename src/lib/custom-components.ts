import type { ContentBlock, InlineNode, TextBlock } from './document'

export type ComponentOverrides = {
  text: string
  url: string
}

export type ComponentBlock = {
  type: 'component'
  runs: []
  componentId: string
  instanceId: string
  initial?: ComponentOverrides
}

export type ComponentDefinition = {
  id: string
  name: string
  kind?: 'button' | 'content'
  marker: string | null
  html: string
  defaults: ComponentOverrides
}

export type ResolvedComponentMarkers = {
  blocks: Array<ContentBlock | ComponentBlock>
  warnings: string[]
}

type ComponentTemplate = Omit<ComponentDefinition, 'defaults'>

// This is trusted application code, never HTML read from a DOCX or user input.
// Keep its markup and inline styles byte-for-byte when changing button fields.
export const ACESSAR_HTML = '<table width="100%" border="0" cellspacing="0" cellpadding="0" role="presentation"><tbody><tr><td align="center"><table border="0" cellspacing="0" cellpadding="0" role="presentation"><tbody><tr><td class="innertd buttonblock" bgcolor="#FD5E04" style="border-radius:7px;-moz-border-radius:7px;-webkit-border-radius:7px;background-color:#FD5E04;"><a target="_blank" class="buttonstyles" style="font-size:16px;font-family:Arial, helvetica, sans-serif;font-weight:bold;color:#FFFFFF;text-align:center;text-decoration:none;display:block;background-color:#FD5E04;border:1px solid transparent;padding:20px;border-radius:7px;-moz-border-radius:7px;-webkit-border-radius:7px;" href="https://pages.xpi.com.br/content/dam/infomoney/TESTE%20BUIDER%20MAIL/JORNADA%20DE%20PFI%20-%20Convers%C3%A3o.docx" title="TEXTO DO BOTÃO" alias="TEXTO DO BOTÃO" conversion="false" data-linkto="https://" data-nl-type="externalLink">Teste doc</a></td></tr></tbody></table></td></tr></tbody></table>'

const definitionsById = new Map<string, ComponentDefinition>()
const definitionsByMarker = new Map<string, ComponentDefinition>()

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function decodeHtml(value: string): string {
  if (typeof document === 'undefined') return value
  const element = document.createElement('textarea')
  element.innerHTML = value
  return element.value
}

function anchorParts(html: string): { open: string; content: string; close: string; start: number; end: number } {
  const match = /(<a\b[^>]*>)([^<]*)(<\/a>)/i.exec(html)
  if (!match) throw new Error('O componente precisa de um link <a> com texto simples para edição segura.')
  return {
    open: match[1],
    content: match[2],
    close: match[3],
    start: match.index,
    end: match.index + match[0].length,
  }
}

function getHref(open: string): string {
  const match = /\bhref=(["'])(.*?)\1/i.exec(open)
  return match ? decodeHtml(match[2]) : ''
}

function safeHttpUrl(value: string): string {
  if (!value || value !== value.trim() || /[\s\u0000-\u001f\u007f]/u.test(value)) {
    throw new Error('Informe uma URL HTTP ou HTTPS válida para o botão.')
  }
  try {
    const parsed = new URL(value)
    const hostname = parsed.hostname.replace(/\.$/, '')
    const qualifiedHost = /^\[[0-9a-f:.]+\]$/i.test(hostname) ||
      (hostname.includes('.') && hostname.split('.').every(Boolean))
    if ((parsed.protocol === 'http:' || parsed.protocol === 'https:') && qualifiedHost) return value
  } catch {
    // A malformed or relative URL cannot be used in an exported email.
  }
  throw new Error('Informe uma URL HTTP ou HTTPS válida para o botão.')
}

function safeButtonUrl(value: string): string {
  // An omitted destination is an intentional, exportable placeholder.
  if (value === '' || value === '#') return '#'
  return safeHttpUrl(value)
}

/** Register a trusted HTML template; marker processing never needs to change. */
export function registerCustomComponent(template: ComponentTemplate): ComponentDefinition {
  if (!/^[a-z][a-z0-9_-]*$/i.test(template.id)) throw new Error('Identificador de componente inválido.')
  if (template.marker !== null && !/^%[a-z][a-z0-9_-]*%$/i.test(template.marker)) {
    throw new Error('Marcador de componente inválido.')
  }
  const marker = template.marker?.toLowerCase()
  if (definitionsById.has(template.id) || (marker && definitionsByMarker.has(marker))) {
    throw new Error(`O componente ${template.id} ou o marcador ${template.marker} já foi registrado.`)
  }

  const anchor = anchorParts(template.html)
  const definition: ComponentDefinition = {
    ...template,
    kind: template.kind ?? 'button',
    defaults: { text: decodeHtml(anchor.content), url: getHref(anchor.open) },
  }
  definitionsById.set(definition.id, definition)
  if (marker) definitionsByMarker.set(marker, definition)
  return definition
}

export function listComponentDefinitions(): ComponentDefinition[] {
  return [...definitionsById.values()]
}

export function getComponentDefinition(id: string): ComponentDefinition | undefined {
  return definitionsById.get(id)
}

function renderTemplate(componentId: string, overrides: Partial<ComponentOverrides>): string {
  const definition = getComponentDefinition(componentId)
  if (!definition) throw new Error(`Componente desconhecido: ${componentId}`)

  const text = overrides.text ?? definition.defaults.text
  const requestedUrl = overrides.url ?? definition.defaults.url
  const url = safeButtonUrl(requestedUrl)
  const anchor = anchorParts(definition.html)
  if (text === definition.defaults.text && url === definition.defaults.url) return definition.html

  const hrefPattern = /(\bhref=)(["'])(.*?)\2/i
  const open = url === definition.defaults.url
    ? anchor.open
    : hrefPattern.test(anchor.open)
      ? anchor.open.replace(hrefPattern, (_match, attribute: string, quote: string) => `${attribute}${quote}${escapeHtml(url)}${quote}`)
      : anchor.open.replace(/>$/, () => ` href="${escapeHtml(url)}">`)
  const content = text === definition.defaults.text ? anchor.content : escapeHtml(text)

  return definition.html.slice(0, anchor.start) + open + content + anchor.close + definition.html.slice(anchor.end)
}

/** Render the original HTML, changing only the first anchor's href and visible text. */
export function renderCustomComponent(componentId: string, overrides: Partial<ComponentOverrides> = {}): string {
  return renderTemplate(componentId, overrides)
}

/** Preview uses the same safe, exportable HTML as the final email. */
export function renderCustomComponentDraft(componentId: string, overrides: Partial<ComponentOverrides> = {}): string {
  return renderTemplate(componentId, overrides)
}

function markerUrl(block: Exclude<ContentBlock, ComponentBlock>, start: number, end: number): { url: string; warning?: string } {
  const urls = new Set<string>()
  let offset = 0
  for (const run of block.runs) {
    const next = offset + run.text.length
    if (run.href && next > start && offset < end) urls.add(run.href)
    offset = next
  }
  if (urls.size > 1) return { url: '#', warning: 'O marcador possui links diferentes; o botão usará # até você definir a URL nas configurações.' }
  const [candidate] = urls
  if (!candidate) return { url: '#', warning: 'O marcador não possui hyperlink no Word; o botão usará # até você definir a URL nas configurações.' }
  if (candidate === '#') return { url: '#' }
  if (candidate.length > 4_096) return { url: '#', warning: 'O hyperlink do marcador é muito longo; o botão usará # até você definir a URL nas configurações.' }
  try {
    return { url: safeHttpUrl(candidate) }
  } catch {
    return { url: '#', warning: 'O hyperlink do marcador não é uma URL HTTP ou HTTPS válida; o botão usará # até você definir a URL nas configurações.' }
  }
}

type ParagraphLine = { start: number; end: number; after: number; before: number }

function paragraphLines(text: string): ParagraphLine[] {
  const lines: ParagraphLine[] = []
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

/** Replace standalone paragraph markers, including markers split across Word runs. */
export function resolveComponentMarkers(blocks: ContentBlock[]): ResolvedComponentMarkers {
  const resolved: ResolvedComponentMarkers['blocks'] = []
  const warnings: string[] = []
  const counts = new Map<string, number>()
  const usedIds = new Set<string>()

  // Reopening an older project may contain both resolved components and raw markers.
  // Reserve every existing ID before allocating new ones, regardless of block order.
  for (const block of blocks) {
    if (block.type !== 'component') continue
    usedIds.add(block.instanceId)
    const prefix = `${block.componentId}-`
    if (block.instanceId.startsWith(prefix)) {
      const suffix = block.instanceId.slice(prefix.length)
      if (/^[1-9]\d*$/.test(suffix)) {
        const numericSuffix = Number(suffix)
        if (Number.isSafeInteger(numericSuffix)) {
          counts.set(block.componentId, Math.max(counts.get(block.componentId) ?? 0, numericSuffix))
        }
      }
    }
  }

  function nextId(componentId: string): string {
    let count = counts.get(componentId) ?? 0
    let instanceId: string
    do {
      count += 1
      instanceId = `${componentId}-${count}`
    } while (usedIds.has(instanceId))
    counts.set(componentId, count)
    usedIds.add(instanceId)
    return instanceId
  }

  function componentForLine(block: TextBlock, text: string, line: ParagraphLine): ComponentBlock | null {
    const lineText = text.slice(line.start, line.end)
    const trimmed = lineText.trim()
    const explicit = definitionsByMarker.get(trimmed.toLowerCase())
    if (explicit) {
      return { type: 'component', runs: [], componentId: explicit.id, instanceId: nextId(explicit.id) }
    }

    const genericMarker = /^%([^%]+)%$/u.exec(trimmed)
    if (!genericMarker) return null
    const label = genericMarker[1].trim()
    if (!label) return null
    if (label.length > 1_000 || /[\u0000-\u001f\u007f-\u009f]/u.test(label)) {
      warnings.push('Um marcador entre % foi preservado como texto: o nome do botão excede 1.000 caracteres ou contém caracteres de controle.')
      return null
    }

    const start = line.start + lineText.length - lineText.trimStart().length
    const link = markerUrl(block, start, start + trimmed.length)
    if (link.warning) warnings.push(`${trimmed}: ${link.warning}`)
    return {
      type: 'component', runs: [], componentId: 'word-button', instanceId: nextId('word-button'),
      initial: { text: label, url: link.url },
    }
  }

  function warnInlineMarkers(text: string): void {
    for (const match of text.matchAll(/%[^%\r\n\u2028\u2029]+%/gu)) {
      warnings.push(`O marcador ${match[0]} precisa ficar sozinho em uma linha do parágrafo para virar botão; o texto original foi preservado.`)
    }
  }

  for (const block of blocks) {
    if (block.type === 'component') {
      resolved.push(block)
      continue
    }
    const text = block.runs.map((run) => run.text).join('')
    if (block.type !== 'paragraph') {
      resolved.push(block)
      warnInlineMarkers(text)
      continue
    }

    const lines = paragraphLines(text)
    const converted: Array<{ line: ParagraphLine; component: ComponentBlock }> = []
    for (const line of lines) {
      const warningCount = warnings.length
      const component = componentForLine(block, text, line)
      if (component) converted.push({ line, component })
      else if (warnings.length === warningCount) warnInlineMarkers(text.slice(line.start, line.end))
    }
    if (converted.length === 0) {
      resolved.push(block)
      continue
    }

    let cursor = 0
    for (const { line, component } of converted) {
      const end = line.start === 0 ? 0 : line.before
      if (end > cursor) {
        resolved.push({ ...block, runs: sliceRuns(block.runs, cursor, end) })
      }
      resolved.push(component)
      cursor = line.after
    }
    if (cursor < text.length) {
      resolved.push({ ...block, runs: sliceRuns(block.runs, cursor, text.length) })
    }
  }

  return { blocks: resolved, warnings: [...new Set(warnings)] }
}

registerCustomComponent({
  id: 'acessar',
  name: 'Botão Acessar',
  marker: '%acessar%',
  html: ACESSAR_HTML,
})

// Generic Word markers reuse the exact trusted button markup, but their text and
// destination are captured separately for each occurrence in the DOCX.
registerCustomComponent({
  id: 'word-button',
  name: 'Botão do Word',
  marker: null,
  html: ACESSAR_HTML,
})
