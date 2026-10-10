import DOMPurify from 'dompurify';
// Mammoth's published browser bundle supports ArrayBuffer input; its package entry
// resolves to the Node ZIP reader under Vitest despite jsdom's browser globals.
// @ts-ignore -- Mammoth does not publish declarations for this browser subpath.
import * as mammothBrowser from 'mammoth/mammoth.browser.js';
import { resolveComponentMarkers } from './custom-components';
import type { ComponentBlock } from './custom-components';
import { resolveConditionMarkers } from './conditions';
import type { ConditionBoundaryBlock } from './conditions';

const mammoth = mammothBrowser as typeof import('mammoth');

export type InlineNode = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  href?: string;
};

export type TextBlock = {
  type: 'paragraph' | 'heading' | 'list-item';
  runs: InlineNode[];
  level?: number;
  ordered?: boolean;
};

export type ContentBlock = TextBlock | ComponentBlock | ConditionBoundaryBlock;

export type ExtractedDocument = {
  blocks: ContentBlock[];
  warnings: string[];
  errors?: string[];
  sourceName: string;
};

type ParsedContent = Pick<ExtractedDocument, 'blocks' | 'warnings' | 'errors'>;
type InlineMarks = Pick<InlineNode, 'bold' | 'italic' | 'href'>;

const ALLOWED_TAGS = [
  'a', 'b', 'blockquote', 'br', 'div', 'em', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
  'i', 'li', 'ol', 'p', 'span', 'strong', 'table', 'tbody', 'td', 'th', 'thead',
  'tr', 'ul',
];

function safeHref(value: string): string | undefined {
  const href = value.trim();
  if (!href) return undefined;

  if (/^#[A-Za-z0-9._-]+$/.test(href)) return href;

  try {
    const url = new URL(href);
    if ((url.protocol === 'http:' || url.protocol === 'https:') && url.hostname) {
      return href;
    }
    if ((url.protocol === 'mailto:' || url.protocol === 'tel:') && url.pathname) {
      return href;
    }
  } catch {
    // Relative and malformed links cannot be used reliably in an exported email.
  }

  return undefined;
}

function appendRun(runs: InlineNode[], text: string, marks: InlineMarks): void {
  if (!text) return;

  const previous = runs[runs.length - 1];
  if (previous && previous.bold === marks.bold && previous.italic === marks.italic && previous.href === marks.href) {
    previous.text += text;
    return;
  }

  runs.push({ text, ...marks });
}

function readInline(node: Node, runs: InlineNode[], marks: InlineMarks, warnings: string[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    appendRun(runs, node.textContent ?? '', marks);
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;

  const element = node as Element;
  const tag = element.tagName.toLowerCase();
  if (tag === 'br') {
    appendRun(runs, '\n', marks);
    return;
  }

  const nextMarks: InlineMarks = { ...marks };
  if (tag === 'b' || tag === 'strong') nextMarks.bold = true;
  if (tag === 'i' || tag === 'em') nextMarks.italic = true;
  if (tag === 'a') {
    const rawHref = element.getAttribute('href');
    if (rawHref) {
      const href = safeHref(rawHref);
      if (href) nextMarks.href = href;
      else warnings.push('Um link com endereço inválido ou não permitido foi mantido apenas como texto.');
    }
  }

  for (const child of element.childNodes) readInline(child, runs, nextMarks, warnings);
}

function runsFrom(nodes: Iterable<Node>, warnings: string[]): InlineNode[] {
  const runs: InlineNode[] = [];
  for (const node of nodes) readInline(node, runs, {}, warnings);
  return runs;
}

function parseBlocks(root: ParentNode, warnings: string[]): TextBlock[] {
  const blocks: TextBlock[] = [];

  function visit(node: Node, list?: { ordered: boolean; level: number }): void {
    if (node.nodeType === Node.TEXT_NODE) {
      if (node.textContent?.trim()) blocks.push({ type: 'paragraph', runs: runsFrom([node], warnings) });
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;

    const element = node as Element;
    const tag = element.tagName.toLowerCase();

    if (tag === 'ul' || tag === 'ol') {
      const nextList = { ordered: tag === 'ol', level: (list?.level ?? 0) + 1 };
      for (const child of element.childNodes) visit(child, nextList);
      return;
    }

    if (tag === 'li') {
      const inlineChildren = Array.from(element.childNodes).filter((child) => {
        if (child.nodeType !== Node.ELEMENT_NODE) return true;
        const childTag = (child as Element).tagName.toLowerCase();
        return childTag !== 'ul' && childTag !== 'ol';
      });
      blocks.push({
        type: 'list-item',
        runs: runsFrom(inlineChildren, warnings),
        ordered: list?.ordered ?? false,
        level: list?.level ?? 1,
      });
      for (const child of element.childNodes) {
        if (child.nodeType === Node.ELEMENT_NODE && /^(ul|ol)$/i.test((child as Element).tagName)) {
          visit(child, list);
        }
      }
      return;
    }

    if (tag === 'p' || /^h[1-6]$/.test(tag)) {
      blocks.push({
        type: tag === 'p' ? 'paragraph' : 'heading',
        runs: runsFrom(element.childNodes, warnings),
        ...(tag === 'p' ? {} : { level: Number(tag.slice(1)) }),
      });
      return;
    }

    // Tables and wrappers are walked in document order. Their text is not lost.
    const hasBlockChildren = Array.from(element.children).some((child) =>
      /^(p|h[1-6]|ul|ol|li|div|blockquote|table|thead|tbody|tr|td|th)$/i.test(child.tagName),
    );
    if (hasBlockChildren) {
      for (const child of element.childNodes) visit(child, list);
    } else {
      blocks.push({ type: 'paragraph', runs: runsFrom([element], warnings) });
    }
  }

  for (const child of root.childNodes) visit(child);
  return blocks;
}

/** Converts Mammoth's HTML into a small, safe document model for preview and email export. */
export function parseMammothHtml(html: string): ParsedContent {
  const warnings: string[] = [];
  if (/<img(?:\s|\/|>)/i.test(html)) {
    warnings.push('Imagens presentes no Word foram ignoradas nesta etapa.');
  }

  const clean = DOMPurify.sanitize(html, {
    ALLOWED_TAGS,
    ALLOWED_ATTR: ['href'],
    ALLOW_ARIA_ATTR: false,
    ALLOW_DATA_ATTR: false,
    ALLOW_UNKNOWN_PROTOCOLS: false,
  });

  if (DOMPurify.removed.some((item) => 'attribute' in item && item.attribute?.name.toLowerCase() === 'href')) {
    warnings.push('Um link com protocolo não permitido foi mantido apenas como texto.');
  }

  const document = new DOMParser().parseFromString(clean, 'text/html');
  const components = resolveComponentMarkers(parseBlocks(document.body, warnings));
  const conditions = resolveConditionMarkers(components.blocks);
  return {
    blocks: conditions.blocks,
    warnings: [...new Set([...warnings, ...components.warnings])],
    errors: conditions.errors,
  };
}

/** Reads a local DOCX without sending its contents to a server. */
export async function extractDocx(file: File): Promise<ExtractedDocument> {
  if (!/\.docx$/i.test(file.name)) throw new Error('Selecione um arquivo .docx.');
  if (file.size === 0) throw new Error('O arquivo DOCX está vazio.');

  const result = await mammoth.convertToHtml(
    { arrayBuffer: await file.arrayBuffer() },
    { externalFileAccess: false, includeEmbeddedStyleMap: false, ignoreEmptyParagraphs: false },
  );
  const parsed = parseMammothHtml(result.value);
  return {
    ...parsed,
    sourceName: file.name,
    warnings: [...result.messages.map((message) => message.message), ...parsed.warnings],
  };
}
