import type { ContentBlock } from './document'
import { getComponentDefinition } from './custom-components'

export type EmailSectionName = 'CABEÇALHO' | 'PARTE FINAL' | `PARTE ${number}` | `BOTÃO ${number}`

/** Indices refer to the complete body block list; end is exclusive. */
export type EmailSectionSegment = {
  name: EmailSectionName
  kind: 'part' | 'button'
  start: number
  end: number
}

export type EmailSectionPlan = {
  buttonCount: number
  segments: EmailSectionSegment[]
}

export function sectionComment(name: EmailSectionName, edge: 'start' | 'end'): string {
  return `<!-- ${name} - ${edge === 'start' ? 'INÍCIO' : 'FIM'} -->`
}

/**
 * Plan comments from the component model before rendering or audience simulation.
 * Button identity comes from its registered definition. Text links and content
 * components remain ordinary body blocks, even with the same destination.
 *
 * A condition boundary ends the current segment. A following segment may retain
 * its logical part name, so no comment pair spans an Adobe IF opening or closing.
 * Blank-only parts receive no markers; their original blocks still render.
 */
export function planEmailSections(blocks: readonly ContentBlock[]): EmailSectionPlan {
  const isButton = (block: ContentBlock): boolean =>
    block.type === 'component' && getComponentDefinition(block.componentId)?.kind === 'button'
  const buttonCount = blocks.reduce((count, block) => count + Number(isButton(block)), 0)
  const segments: EmailSectionSegment[] = []
  let seenButtons = 0
  let partStart: number | null = null
  let partHasContent = false

  function finishPart(end: number): void {
    if (partStart === null) return
    if (partHasContent) {
      segments.push({
        name: buttonCount > 0 && seenButtons === buttonCount ? 'PARTE FINAL' : `PARTE ${seenButtons + 1}`,
        kind: 'part',
        start: partStart,
        end,
      })
    }
    partStart = null
    partHasContent = false
  }

  blocks.forEach((block, index) => {
    if (block.type === 'condition-start' || block.type === 'condition-end') {
      finishPart(index)
      return
    }
    if (isButton(block)) {
      finishPart(index)
      seenButtons += 1
      segments.push({ name: `BOTÃO ${seenButtons}`, kind: 'button', start: index, end: index + 1 })
      return
    }
    partStart ??= index
    partHasContent ||= block.type === 'component' || block.runs.some((run) => run.text.trim().length > 0)
  })

  finishPart(blocks.length)
  return { buttonCount, segments }
}
