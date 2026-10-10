import assert from 'node:assert/strict'
import { File } from 'node:buffer'
import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { JSDOM } from 'jsdom'
import { createServer } from 'vite'
import mammothBrowser from 'mammoth/mammoth.browser.js'

// Reproduce the Word -> document model -> real email generator pipeline locally.
// Fixed ZIP/project dates keep the sample artifacts reproducible.
const projectRoot = fileURLToPath(new URL('../', import.meta.url))
const outputDirectory = fileURLToPath(new URL('../examples/', import.meta.url))
const basename = 'organizacao-quatro-botoes'
const fixedDate = new Date('2026-10-10T00:00:00.000Z')
const xmlEscape = (text) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
const run = (text, marks = '') => `<w:r>${marks ? `<w:rPr>${marks}</w:rPr>` : ''}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`
const paragraph = (...runs) => `<w:p>${runs.join('')}</w:p>`
const linkedButton = (label, id) => paragraph(run('%'), `<w:hyperlink r:id="${id}">${run(label, '<w:b/>')}</w:hyperlink>`, run('%'))

const zip = new JSZip()
const addXml = (name, contents) => zip.file(name, contents, { date: fixedDate })
addXml('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
addXml('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)
const links = [
  ['rIdInfo', 'https://example.org/informacoes'],
  ['rIdButton1', 'https://example.org/novidades'],
  ['rIdButton2', 'https://example.org/base-ativo'],
  ['rIdButton3', 'https://example.org/clientes-vip'],
  ['rIdButton4', 'https://example.org/proximos-passos'],
]
addXml('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${links.map(([id, url]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlEscape(url)}" TargetMode="External"/>`).join('\n  ')}
</Relationships>`)
const paragraphs = [
  paragraph(run('Assunto: '), run('Novidades para você — campanha de exemplo', '<w:b/>')),
  paragraph(run('Pré-header: '), run('Confira quatro oportunidades e os conteúdos de cada audiência.', '<w:i/>')),
  paragraph(run('Olá, cliente!')),
  paragraph(run('Preparamos uma '), run('seleção de novidades', '<w:b/>'), run(' para você. '), `<w:hyperlink r:id="rIdInfo">${run('Leia também as informações gerais.')}</w:hyperlink>`),
  linkedButton('CONHECER NOVIDADES', 'rIdButton1'),
  paragraph(run('Depois das novidades, confira os conteúdos disponíveis para cada perfil.')),
  paragraph(run('##BASE'), run('_ATIVO##', '<w:b/>')),
  paragraph(run('Você possui acesso a '), run('conteúdos exclusivos', '<w:b/>'), run(' para a audiência BASE_ATIVO.')),
  linkedButton('ACESSAR CONTEÚDO EXCLUSIVO', 'rIdButton2'),
  paragraph(run('A seleção exclusiva permanece disponível para consulta.', '<w:i/>')),
  paragraph(run('##FIM_BASE_ATIVO##')),
  paragraph(run('Conheça também nossas recomendações para acompanhar as próximas ações.')),
  paragraph(run('##CLIENTES_VIP##')),
  paragraph(run('Clientes da audiência CLIENTES_VIP podem conferir os benefícios especiais.')),
  linkedButton('VER BENEFÍCIOS VIP', 'rIdButton3'),
  paragraph(run('Consulte as orientações antes de escolher seu benefício.')),
  paragraph(run('##FIM_CLIENTES_'), run('VIP##', '<w:i/>')),
  paragraph(run('Para continuar, veja os próximos passos desta campanha de exemplo.')),
  linkedButton('ABRIR PRÓXIMOS PASSOS', 'rIdButton4'),
  paragraph(run('Obrigado por acompanhar nossas novidades!')),
  paragraph(run('Até breve,')),
  paragraph(run('Equipe Exemplo', '<w:b/>')),
]
addXml('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    ${paragraphs.join('\n    ')}
    <w:sectPr/>
  </w:body>
</w:document>`)
// JSZip also creates directory entries; fix their dates as well as the XML files.
for (const entry of Object.values(zip.files)) entry.date = fixedDate
const docx = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })

const dom = new JSDOM('<!doctype html><html><body></body></html>', { url: 'http://localhost/' })
for (const key of ['window', 'document', 'DOMParser', 'Node', 'Element', 'HTMLElement']) {
  globalThis[key] = dom.window[key]
}
globalThis.self = dom.window
globalThis.__ATELIE_EXAMPLE_MAMMOTH__ = mammothBrowser
globalThis.__ATELIE_EXAMPLE_MJML__ = (await import('mjml-browser')).default
const server = await createServer({
  root: projectRoot,
  configFile: false,
  plugins: [{
    name: 'example-browser-bundle-interop',
    enforce: 'pre',
    transform(code, id) {
      // Vite's browser build exposes CJS named exports automatically. Its SSR
      // loader needs the browser bundle loaded through Node's native CJS
      // interop. This adaptation applies in memory; source stays unchanged.
      if (id.endsWith('/src/lib/document.ts')) {
        return code.replace(
          "import * as mammothBrowser from 'mammoth/mammoth.browser.js';",
          'const mammothBrowser = globalThis.__ATELIE_EXAMPLE_MAMMOTH__;',
        )
      }
      if (id.endsWith('/src/lib/email.ts')) {
        return code.replace(
          "const { default: mjml2html } = await import('mjml-browser')",
          'const mjml2html = globalThis.__ATELIE_EXAMPLE_MJML__',
        )
      }
    },
  }],
  server: { middlewareMode: true, ws: false },
  appType: 'custom',
  logLevel: 'error',
})
try {
  const { extractDocx } = await server.ssrLoadModule('/src/lib/document.ts')
  const { renderEmail } = await server.ssrLoadModule('/src/lib/email.ts')
  const { DEFAULT_HEADER_SETTINGS } = await server.ssrLoadModule('/src/lib/header.ts')
  const { renderCustomComponent } = await server.ssrLoadModule('/src/lib/custom-components.ts')
  const { createProjectData, serializeProject, parseProject } = await server.ssrLoadModule('/src/lib/project.ts')
  const document = await extractDocx(new File([docx], `${basename}.docx`, {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  }))
  assert.deepEqual(document.errors, [])
  assert.deepEqual(document.warnings, [])
  const buttons = document.blocks.filter((block) => block.type === 'component')
  assert.equal(buttons.length, 4)
  assert.deepEqual(buttons.map((button) => button.initial.url), links.slice(1).map(([, url]) => url))
  const settings = {
    width: 600,
    fontFamily: 'Arial, Helvetica, sans-serif',
    backgroundColor: '#f3f4f6',
    surfaceColor: '#ffffff',
    textColor: '#1f2937',
    accentColor: '#2563eb',
    template: 'classic',
    header: { ...DEFAULT_HEADER_SETTINGS },
  }
  const project = createProjectData(document, {}, settings)
  project.savedAt = fixedDate.toISOString()
  const projectJson = JSON.stringify(JSON.parse(serializeProject(project)), null, 2) + '\n'
  assert.equal(parseProject(projectJson).document.blocks.length, document.blocks.length)
  const result = await renderEmail(document, settings)
  assert.deepEqual(result.errors, [])
  assert.ok(result.html.indexOf('Pré-header:') < result.html.indexOf('bgcolor="#000000"'))
  assert.deepEqual([...result.html.matchAll(/<!-- BOTÃO (\d+) - INÍCIO -->/gu)].map((match) => match[1]), ['1', '2', '3', '4'])
  for (const [index, audience] of ['BASE_ATIVO', 'CLIENTES_VIP'].entries()) {
    const opening = `{%#if profile.segmentMembership.get("ups").get("${audience}").status = "existing"%}`
    const start = result.html.indexOf(opening)
    const end = result.html.indexOf('{%/if%}', start)
    assert.ok(start >= 0 && end > start)
    assert.ok(result.html.slice(start, end).includes(`<!-- BOTÃO ${index + 2} - INÍCIO -->`))
    assert.ok(result.html.slice(start, end).includes(`<!-- BOTÃO ${index + 2} - FIM -->`))
  }
  for (const button of buttons) assert.ok(result.html.includes(renderCustomComponent(button.componentId, button.initial)))
  assert.equal(result.html.split('{%/if%}').length - 1, 2)
  assert.ok(result.html.includes('<!-- PARTE FINAL - INÍCIO -->'))
  await mkdir(outputDirectory, { recursive: true })
  for (const [suffix, content] of [
    ['.docx', docx], ['.json', projectJson], ['-email.html', result.html],
  ]) {
    const path = `${outputDirectory}${basename}${suffix}`
    await writeFile(path, content)
    console.log(path)
  }
} finally {
  await server.close()
  dom.window.close()
  delete globalThis.__ATELIE_EXAMPLE_MAMMOTH__
  delete globalThis.__ATELIE_EXAMPLE_MJML__
}
