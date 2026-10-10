import { mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'

const zip = new JSZip()
zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`)
zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdInfo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/informacoes" TargetMode="External"/>
</Relationships>`)
zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>SUB: Exemplo de email</w:t></w:r></w:p>
    <w:p><w:r><w:t>Linha fina: veja as condições de audiência</w:t></w:r></w:p>
    <w:p><w:r><w:t>Olá, cliente!</w:t></w:r></w:p>
    <w:p><w:r><w:t>##BASE</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>_ATIVO##</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Saber quanto deseja receber em </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>renda mensal</w:t></w:r><w:r><w:t xml:space="preserve"> é fundamental. </w:t></w:r><w:hyperlink r:id="rIdInfo"><w:r><w:t>Veja os detalhes.</w:t></w:r></w:hyperlink></w:p>
    <w:p><w:r><w:t>%acessar%</w:t></w:r></w:p>
    <w:p><w:r><w:t>##FIM_BASE_</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ATIVO##</w:t></w:r></w:p>
    <w:p><w:r><w:t>Conteúdo para todos.</w:t></w:r></w:p>
    <w:p><w:r><w:t>##CLIENTES_VIP##</w:t></w:r></w:p>
    <w:p><w:r><w:t>Uma oferta especial para clientes VIP.</w:t></w:r></w:p>
    <w:p><w:r><w:t>%BAIXAR AGORA%</w:t></w:r></w:p>
    <w:p><w:r><w:t>##FIM_CLIENTES_VIP##</w:t></w:r></w:p>
    <w:p><w:r><w:t>Até breve!</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)

const output = fileURLToPath(new URL('../examples/condicoes-audiencia.docx', import.meta.url))
await mkdir(fileURLToPath(new URL('../examples/', import.meta.url)), { recursive: true })
await writeFile(output, await zip.generateAsync({ type: 'nodebuffer' }))
console.log(output)
