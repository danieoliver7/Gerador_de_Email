import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { extractDocx, parseMammothHtml } from './document';
import { renderEmail, type EmailSettings } from './email';
import { DEFAULT_HEADER_SETTINGS } from './header';

const emailSettings: EmailSettings = {
  width: 600,
  fontFamily: 'Arial, Helvetica, sans-serif',
  backgroundColor: '#eef2f7',
  surfaceColor: '#ffffff',
  textColor: '#172033',
  accentColor: '#2054d2',
  template: 'classic',
  header: DEFAULT_HEADER_SETTINGS,
};

async function createFixtureDocx(bodyXml = `
          <w:p>
            <w:r><w:t xml:space="preserve">Olá </w:t></w:r>
            <w:r><w:rPr><w:b/></w:rPr><w:t>cliente</w:t></w:r>
            <w:r><w:t xml:space="preserve">, veja </w:t></w:r>
            <w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:i/></w:rPr><w:t>oferta</w:t></w:r></w:hyperlink>
          </w:p>
          <w:p><w:r><w:t>Texto &amp; informação.</w:t></w:r></w:p>
          <w:p/>
  `, linkTarget = 'https://example.org/oferta?a=1&b=2'): Promise<File> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>`,
  );
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${linkTarget.replace(/&/g, '&amp;').replace(/"/g, '&quot;')}" TargetMode="External"/>
      </Relationships>`,
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>${bodyXml}</w:body>
      </w:document>`,
  );

  const arrayBuffer = await zip.generateAsync({ type: 'arraybuffer' });
  const file = new File([arrayBuffer], 'copy.docx', {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  // jsdom's File implementation may not expose Blob.arrayBuffer.
  if (!file.arrayBuffer) Object.defineProperty(file, 'arrayBuffer', { value: async () => arrayBuffer });
  return file;
}

describe('document extraction', () => {
  it('reads a DOCX with ordered paragraphs, emphasis and an original hyperlink', async () => {
    const result = await extractDocx(await createFixtureDocx());

    expect(result.sourceName).toBe('copy.docx');
    expect(result.blocks.map((block) => block.runs.map((run) => run.text).join(''))).toEqual([
      'Olá cliente, veja oferta',
      'Texto & informação.',
      '',
    ]);
    expect(result.blocks[0].runs).toContainEqual({ text: 'cliente', bold: true });
    expect(result.blocks[0].runs).toContainEqual({
      text: 'oferta',
      italic: true,
      href: 'https://example.org/oferta?a=1&b=2',
    });
  });

  it('keeps list order and nesting in the structured output', () => {
    const result = parseMammothHtml('<p>Início</p><ul><li>Um</li><li>Dois<ol><li>Três</li></ol></li></ul>');

    expect(result.blocks.map((block) => ({
      type: block.type,
      text: block.runs.map((run) => run.text).join(''),
      level: 'level' in block ? block.level : undefined,
      ordered: 'ordered' in block ? block.ordered : undefined,
    }))).toEqual([
      { type: 'paragraph', text: 'Início', level: undefined, ordered: undefined },
      { type: 'list-item', text: 'Um', level: 1, ordered: false },
      { type: 'list-item', text: 'Dois', level: 1, ordered: false },
      { type: 'list-item', text: 'Três', level: 2, ordered: true },
    ]);
  });

  it('never retains active HTML, unsafe links or images', () => {
    const result = parseMammothHtml(
      '<p>Antes <a href="javascript:alert(1)">clique</a> <strong>forte</strong></p>' +
      '<img src="https://example.org/track.png"><script>alert(2)</script><p>Depois</p>',
    );

    expect(result.blocks.map((block) => block.runs.map((run) => run.text).join(''))).toEqual([
      'Antes clique forte',
      'Depois',
    ]);
    expect(result.blocks.flatMap((block) => block.runs).every((run) => !run.href)).toBe(true);
    expect(result.warnings).toContain('Imagens presentes no Word foram ignoradas nesta etapa.');
    expect(result.warnings).toContain('Um link com protocolo não permitido foi mantido apenas como texto.');
  });

  it('rejects non-DOCX files before conversion', async () => {
    await expect(extractDocx(new File(['x'], 'copy.doc'))).rejects.toThrow('Selecione um arquivo .docx.');
  });

  it('turns a single standalone marker in a DOCX into a component', async () => {
    const result = await extractDocx(await createFixtureDocx(
      '<w:p><w:r><w:t>%acessar%</w:t></w:r></w:p>',
    ));
    expect(result.blocks).toEqual([{ type: 'component', runs: [], componentId: 'acessar', instanceId: 'acessar-1' }]);
  });

  it('keeps the position of several markers even when Word splits their runs', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>Antes</w:t></w:r></w:p>
      <w:p><w:r><w:t>%ace</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ssar%</w:t></w:r></w:p>
      <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Depois</w:t></w:r></w:p>
      <w:p><w:r><w:t>%acessar%</w:t></w:r></w:p>
    `));
    expect(result.blocks.map((block) => block.type)).toEqual(['paragraph', 'component', 'paragraph', 'component']);
    expect(result.blocks.filter((block) => block.type === 'component').map((block) => block.instanceId))
      .toEqual(['acessar-1', 'acessar-2']);
    expect(result.blocks[2].runs).toEqual([{ text: 'Depois', italic: true }]);
  });

  it('turns a linked %texto% paragraph from a real DOCX into a button with the Word label and URL', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>Antes</w:t></w:r></w:p>
      <w:p>
        <w:r><w:t>%</w:t></w:r>
        <w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">BAIXAR </w:t></w:r><w:r><w:t>AGORA</w:t></w:r></w:hyperlink>
        <w:r><w:t>%</w:t></w:r>
      </w:p>
      <w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Depois</w:t></w:r></w:p>
    `));

    expect(result.blocks).toEqual([
      { type: 'paragraph', runs: [{ text: 'Antes' }] },
      {
        type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
        initial: { text: 'BAIXAR AGORA', url: 'https://example.org/oferta?a=1&b=2' },
      },
      { type: 'paragraph', runs: [{ text: 'Depois', italic: true }] },
    ]);
    expect(result.warnings.join(' ')).not.toMatch(/precisa ficar sozinho/);
  });

  it('recognizes several generic markers split over Word runs and keeps their document order', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>%BAI</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>XAR AGORA%</w:t></w:r></w:p>
      <w:p><w:r><w:t>Entre os botões</w:t></w:r></w:p>
      <w:p><w:r><w:t>%</w:t></w:r><w:r><w:t>SAIBA MAIS</w:t></w:r><w:r><w:t>%</w:t></w:r></w:p>
    `));

    expect(result.blocks.map((block) => block.type)).toEqual(['component', 'paragraph', 'component']);
    expect(result.blocks[0]).toEqual({
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url: '#' },
    });
    expect(result.blocks[2]).toEqual({
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-2',
      initial: { text: 'SAIBA MAIS', url: '#' },
    });
  });

  it('preserves an inline %texto% marker as copy and warns about its position', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t xml:space="preserve">Clique em %BAIXAR </w:t></w:r><w:r><w:t>AGORA% para continuar.</w:t></w:r></w:p>
    `));

    expect(result.blocks).toEqual([
      { type: 'paragraph', runs: [{ text: 'Clique em %BAIXAR AGORA% para continuar.' }] },
    ]);
    expect(result.warnings.join(' ')).toMatch(/%BAIXAR AGORA%/i);
    expect(result.warnings.join(' ')).toMatch(/linha|parágrafo/i);
  });

  it('recognizes a linked marker on its own Word line despite breaks, tabs and NBSP in one paragraph', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Antes</w:t></w:r>
        <w:r><w:br/></w:r>
        <w:r><w:tab/></w:r>
        <w:r><w:t xml:space="preserve">&#160; </w:t></w:r>
        <w:r><w:t>%</w:t></w:r>
        <w:hyperlink r:id="rIdLink"><w:r><w:t xml:space="preserve">BAIXAR </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>AGORA</w:t></w:r></w:hyperlink>
        <w:r><w:t>%</w:t></w:r>
        <w:r><w:br/></w:r>
        <w:r><w:rPr><w:i/></w:rPr><w:t>Depois</w:t></w:r>
      </w:p>
    `));

    expect(result.blocks.map((block) => block.type)).toEqual(['paragraph', 'component', 'paragraph']);
    expect(result.blocks[0].runs.map((run) => run.text).join('')).toBe('Antes');
    expect(result.blocks[0].runs.some((run) => run.bold && run.text.includes('Antes'))).toBe(true);
    expect(result.blocks[1]).toEqual({
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url: 'https://example.org/oferta?a=1&b=2' },
    });
    expect(result.blocks[2].runs.map((run) => run.text).join('')).toBe('Depois');
    expect(result.blocks[2].runs.some((run) => run.italic && run.text.includes('Depois'))).toBe(true);
    expect(result.warnings.join(' ')).not.toMatch(/precisa ficar sozinho/);

    const email = await renderEmail(result, emailSettings);
    const buttonPosition = email.html.indexOf('>BAIXAR AGORA</a>');
    expect(email.errors).toEqual([]);
    expect(email.html.indexOf('<strong>Antes</strong>')).toBeLessThan(buttonPosition);
    expect(buttonPosition).toBeLessThan(email.html.indexOf('<em>Depois</em>'));
    expect(email.html).not.toContain('%BAIXAR AGORA%');
  });

  it('uses # for an incomplete Word hyperlink and still exports the button', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>%</w:t></w:r><w:hyperlink r:id="rIdLink"><w:r><w:t>BAIXAR AGORA</w:t></w:r></w:hyperlink><w:r><w:t>%</w:t></w:r></w:p>
    `, 'https://www/'));

    expect(result.blocks).toEqual([{
      type: 'component', runs: [], componentId: 'word-button', instanceId: 'word-button-1',
      initial: { text: 'BAIXAR AGORA', url: '#' },
    }]);

    const fallback = await renderEmail(result, emailSettings);
    expect(fallback.errors).toEqual([]);
    expect(fallback.html).toContain('href="#"');
    expect(fallback.html).toContain('>BAIXAR AGORA</a>');
    expect(fallback.html).not.toContain('href="https://www/"');

    const configured = await renderEmail(result, emailSettings, {
      'word-button-1': { url: 'https://example.org/download' },
    });
    expect(configured.errors).toEqual([]);
    expect(configured.html).toContain('href="https://example.org/download"');
    expect(configured.html).toContain('>BAIXAR AGORA</a>');
  });

  it('imports fragmented audience markers from a real DOCX and protects a Word button', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>Fora</w:t></w:r></w:p>
      <w:p><w:r><w:t>##BASE</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>_ATIVO##</w:t></w:r></w:p>
      <w:p><w:r><w:t xml:space="preserve">Conteúdo </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>forte</w:t></w:r></w:p>
      <w:p><w:r><w:t>%acessar%</w:t></w:r></w:p>
      <w:p><w:r><w:t>##FIM_BASE_</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>ATIVO##</w:t></w:r></w:p>
      <w:p><w:r><w:t>Depois</w:t></w:r></w:p>
    `));
    expect(result.errors).toEqual([]);
    expect(result.blocks.map((block) => block.type)).toEqual([
      'paragraph', 'condition-start', 'paragraph', 'component', 'condition-end', 'paragraph',
    ]);
    const email = await renderEmail(result, emailSettings);
    expect(email.errors).toEqual([]);
    const opening = email.html.indexOf('{%#if profile.segmentMembership.get("ups").get("BASE_ATIVO").status = "existing"%}');
    const closing = email.html.indexOf('{%/if%}');
    expect(email.html.indexOf('Fora')).toBeLessThan(opening);
    expect(email.html.slice(opening, closing)).toContain('<strong>forte</strong>');
    expect(email.html.slice(opening, closing)).toContain('class="innertd buttonblock"');
    expect(email.html.indexOf('Depois')).toBeGreaterThan(closing);
  });

  it('reports mismatched audience markers in a real DOCX and prevents export', async () => {
    const result = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>##BASE_ATIVO##</w:t></w:r></w:p>
      <w:p><w:r><w:t>Conteúdo</w:t></w:r></w:p>
      <w:p><w:r><w:t>##FIM_CLIENTES_VIP##</w:t></w:r></w:p>
    `));
    expect(result.errors?.join(' ')).toMatch(/não corresponde/);
    const email = await renderEmail(result, emailSettings);
    expect(email.html).toBe('');
    expect(email.errors.join(' ')).toMatch(/não corresponde/);
  });

  it('places the automatic header after SUB and line fine from a real DOCX', async () => {
    const extracted = await extractDocx(await createFixtureDocx(`
      <w:p><w:r><w:t>SUB: prévia do email</w:t></w:r></w:p>
      <w:p><w:r><w:t>Linha fina: separador existente</w:t></w:r></w:p>
      <w:p><w:r><w:t>##BASE_ATIVO##</w:t></w:r></w:p>
      <w:p><w:r><w:t>Corpo para a audiência</w:t></w:r></w:p>
      <w:p><w:r><w:t>##FIM_BASE_ATIVO##</w:t></w:r></w:p>
    `));
    const email = await renderEmail(extracted, emailSettings);
    expect(email.errors).toEqual([]);
    const sub = email.html.indexOf('SUB: prévia do email');
    const line = email.html.indexOf('Linha fina: separador existente');
    const header = email.html.indexOf('bgcolor="#000000" height="100"');
    const condition = email.html.indexOf('{%#if profile.segmentMembership.get("ups").get("BASE_ATIVO").status = "existing"%}');
    const copy = email.html.indexOf('Corpo para a audiência');
    expect(sub).toBeGreaterThan(-1);
    expect(sub).toBeLessThan(line);
    expect(line).toBeLessThan(header);
    expect(header).toBeLessThan(condition);
    expect(condition).toBeLessThan(copy);
  });
});
