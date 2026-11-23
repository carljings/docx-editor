// P2 WPS official-document editing capabilities in Core:
// whole-table border presets (GB/T 7713 three-line / all / none), cell vertical alignment & fill,
// cell splitting, header/footer + official page numbers & cm distances, Chinese watermarks,
// TOC / bookmarks, footnote / endnote / textbox, and clearFormatting.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(bodyInner: string, extraParts?: Record<string, string>): Uint8Array {
  const stylesOverride = extraParts?.['word/styles.xml']
    ? '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    : '';
  const stylesRel = extraParts?.['word/styles.xml']
    ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    : '';
  const entries: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        stylesOverride +
        `</Types>`
    ),
    '_rels/.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL}">` +
        `<Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/>` +
        `</Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL}">${stylesRel}</Relationships>`
    ),
    'word/document.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body>${bodyInner}` +
        `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992"/></w:sectPr>` +
        `</w:body></w:document>`
    ),
  };
  if (extraParts) {
    for (const [k, v] of Object.entries(extraParts)) entries[k] = strToU8(v);
  }
  return zipSync(entries);
}

const HEADING_STYLES_XML =
  `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}">` +
  `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>` +
  `</w:styles>`;

function makeTableXml(rows: number, cols: number, headerRows = 0): string {
  const gridCols = Array.from({ length: cols }, () => '<w:gridCol w:w="3000"/>').join('');
  const trs = Array.from({ length: rows }, (_, r) => {
    const trPr = r < headerRows ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
    const tcs = Array.from(
      { length: cols },
      (_, c) => `<w:tc><w:p><w:r><w:t>R${r + 1}C${c + 1}</w:t></w:r></w:p></w:tc>`
    ).join('');
    return `<w:tr>${trPr}${tcs}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr/><w:tblGrid>${gridCols}</w:tblGrid>${trs}</w:tbl>`;
}

const mounted: DocxEditorInstance[] = [];
afterEach(() => {
  while (mounted.length > 0) mounted.pop()!.destroy();
});

function openEditor(bytes: Uint8Array): { editor: DocxEditorInstance; surface: PaginatedSurface } {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const editor = createDocxEditor({ container: host, document: bytes });
  mounted.push(editor);
  const surface = editor.surface;
  if (!surface) throw new Error('surface not mounted');
  return { editor, surface };
}

function selectPara(surface: PaginatedSurface, index: number, from = 0, to = from): void {
  const id = surface.session.paragraphIds()[index]!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: from },
    head: { paragraphId: id, offset: to },
  });
}

function extractRowXmls(docXml: string): string[] {
  return [...docXml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((m) => m[0]!);
}

function hasBorder(
  xml: string,
  side: 'top' | 'bottom' | 'left' | 'right',
  val: string,
  sz?: number
): boolean {
  const re = new RegExp(`<w:${side}\\b[^>]*\\/>`, 'g');
  return [...xml.matchAll(re)].some((m) => {
    const tag = m[0]!;
    if (!tag.includes(`w:val="${val}"`)) return false;
    return sz === undefined ? true : tag.includes(`w:sz="${sz}"`);
  });
}

describe('P2 Whole-table border presets & advanced table ops', () => {
  test('threeLine applies GB/T 7713 borders across whole table from any cell and undoes in one step', () => {
    const { editor, surface } = openEditor(docx(makeTableXml(3, 2)));
    // Place caret in bottom-right body cell (paragraph index 5)
    selectPara(surface, 5, 0, 0);
    const beforeXml = serializeOoxmlPart(surface.session.part());

    const res = editor.exec({ type: 'setTableBorders', scope: 'threeLine' });
    expect(res.ok).toBe(true);

    const rows = extractRowXmls(serializeOoxmlPart(surface.session.part()));
    expect(rows.length).toBe(3);
    // Header row: 1.5pt (sz=12) top, 0.75pt (sz=6) bottom, no left/right
    expect(hasBorder(rows[0]!, 'top', 'single', 12)).toBe(true);
    expect(hasBorder(rows[0]!, 'bottom', 'single', 6)).toBe(true);
    expect(hasBorder(rows[0]!, 'left', 'none')).toBe(true);
    expect(hasBorder(rows[0]!, 'right', 'none')).toBe(true);
    // Middle row: 0.75pt (sz=6) top matching header bottom, none bottom/left/right
    expect(hasBorder(rows[1]!, 'top', 'single', 6)).toBe(true);
    expect(hasBorder(rows[1]!, 'bottom', 'none')).toBe(true);
    // Last row: none top, 1.5pt (sz=12) bottom
    expect(hasBorder(rows[2]!, 'top', 'none')).toBe(true);
    expect(hasBorder(rows[2]!, 'bottom', 'single', 12)).toBe(true);

    // Single undo restores the original table bit-for-bit
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).toBe(beforeXml);
  });

  test('threeLine respects multi-row tblHeader, single-row tables, and vertically merged headers', () => {
    // 1. Multi-row header (2 header rows out of 4)
    const multi = openEditor(docx(makeTableXml(4, 2, 2)));
    selectPara(multi.surface, 6, 0, 0);
    expect(multi.editor.exec({ type: 'setTableBorders', scope: 'threeLine' }).ok).toBe(true);
    const multiRows = extractRowXmls(serializeOoxmlPart(multi.surface.session.part()));
    expect(hasBorder(multiRows[0]!, 'top', 'single', 12)).toBe(true);
    expect(hasBorder(multiRows[0]!, 'bottom', 'none')).toBe(true);
    expect(hasBorder(multiRows[1]!, 'bottom', 'single', 6)).toBe(true);
    expect(hasBorder(multiRows[2]!, 'top', 'single', 6)).toBe(true);
    expect(hasBorder(multiRows[3]!, 'bottom', 'single', 12)).toBe(true);

    // 2. Single-row table
    const single = openEditor(docx(makeTableXml(1, 2)));
    selectPara(single.surface, 0, 0, 0);
    expect(single.editor.exec({ type: 'setTableBorders', scope: 'threeLine' }).ok).toBe(true);
    const singleRows = extractRowXmls(serializeOoxmlPart(single.surface.session.part()));
    expect(hasBorder(singleRows[0]!, 'top', 'single', 12)).toBe(true);
    expect(hasBorder(singleRows[0]!, 'bottom', 'single', 12)).toBe(true);

    // 3. Vertically merged cell spanning rows 0-1 in a 3x2 table
    const mergedXml =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>合并头</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>子列1</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>` +
      `<w:tc><w:p><w:r><w:t>子列2</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>数1</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:p><w:r><w:t>数2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const merged = openEditor(docx(mergedXml));
    selectPara(merged.surface, 4, 0, 0);
    expect(merged.editor.exec({ type: 'setTableBorders', scope: 'threeLine' }).ok).toBe(true);
    const mergedRows = extractRowXmls(serializeOoxmlPart(merged.surface.session.part()));
    expect(hasBorder(mergedRows[0]!, 'top', 'single', 12)).toBe(true);
    expect(hasBorder(mergedRows[2]!, 'top', 'single', 6)).toBe(true);
    expect(hasBorder(mergedRows[2]!, 'bottom', 'single', 12)).toBe(true);
  });

  test('tableAll, tableNone, splitCell, vertical alignment, and cell fill round-trip cleanly', async () => {
    const { editor, surface } = openEditor(docx(makeTableXml(2, 2)));
    selectPara(surface, 0, 0, 0);

    expect(editor.exec({ type: 'setTableBorders', scope: 'tableAll' }).ok).toBe(true);
    const afterAll = serializeOoxmlPart(surface.session.part());
    expect(hasBorder(afterAll, 'left', 'single', 4)).toBe(true);
    expect(hasBorder(afterAll, 'right', 'single', 4)).toBe(true);

    expect(editor.exec({ type: 'setTableBorders', scope: 'tableNone' }).ok).toBe(true);
    const afterNone = serializeOoxmlPart(surface.session.part());
    expect(afterNone).not.toContain('w:val="single"');

    expect(editor.exec({ type: 'setTableCellVerticalAlignment', alignment: 'center' }).ok).toBe(
      true
    );
    expect(editor.exec({ type: 'setCellFill', color: { kind: 'hex', value: 'EFF6FF' } }).ok).toBe(
      true
    );
    expect(editor.exec({ type: 'splitCell', rows: 1, cols: 2 }).ok).toBe(true);

    const saved = unzipSync(new Uint8Array(await editor.save()));
    const docXml = strFromU8(saved['word/document.xml']!);
    expect(docXml).toContain('<w:vAlign w:val="center"');
    expect(docXml).toContain('w:fill="EFF6FF"');
    expect((docXml.match(/<w:tc\b/g) ?? []).length).toBe(5);

    // Clear cell fill on whole table (splitCell cloned fill to the new right half)
    expect(editor.exec({ type: 'selectTableRegion', region: 'table' }).ok).toBe(true);
    expect(editor.exec({ type: 'setCellFill', color: null }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).not.toContain('w:fill="EFF6FF"');
  });

  test('whole-table border presets refuse outside table and reject invalid specs', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>正文</w:t></w:r></w:p>'));
    selectPara(surface, 0, 0, 0);
    expect(editor.exec({ type: 'setTableBorders', scope: 'threeLine' }).ok).toBe(false);

    const inTbl = openEditor(docx(makeTableXml(2, 2)));
    selectPara(inTbl.surface, 0, 0, 0);
    const bad = inTbl.editor.exec({
      type: 'setTableBorders',
      scope: 'threeLine',
      spec: { style: 'single', size: -5, color: { kind: 'hex', value: '000000' } },
    });
    expect(bad.ok).toBe(false);
  });
});

describe('P2 Header/Footer, Official Page Numbers, Watermark, Outline/TOC/Bookmarks & Notes', () => {
  test('official page number footer, titlePage, and cm distances serialize cleanly', async () => {
    const { editor } = openEditor(docx('<w:p><w:r><w:t>公文正文内容</w:t></w:r></w:p>'));

    expect(editor.exec({ type: 'editHeaderFooter', position: 'footer' }).ok).toBe(true);
    expect(editor.exec({ type: 'setAlignment', align: 'center' }).ok).toBe(true);
    expect(editor.exec({ type: 'replaceText', text: '— ' }).ok).toBe(true);
    expect(editor.exec({ type: 'insertPageField', field: 'PAGE' }).ok).toBe(true);
    expect(editor.exec({ type: 'replaceText', text: ' —' }).ok).toBe(true);
    expect(editor.exec({ type: 'exitHeaderFooter' }).ok).toBe(true);

    expect(
      editor.exec({
        type: 'setHeaderFooterOptions',
        titlePage: true,
        headerDistanceTwips: 850,
        footerDistanceTwips: 992,
      }).ok
    ).toBe(true);

    const saved = unzipSync(new Uint8Array(await editor.save()));
    const docXml = strFromU8(saved['word/document.xml']!);
    expect(docXml).toContain('<w:titlePg');
    expect(docXml).toContain('w:header="850"');
    expect(docXml).toContain('w:footer="992"');

    const footerEntries = Object.keys(saved).filter((k) => /^word\/footer\d+\.xml$/.test(k));
    expect(footerEntries.length).toBeGreaterThan(0);
    const footerXml = strFromU8(saved[footerEntries[0]!]!);
    expect(footerXml).toContain('PAGE');
    expect(footerXml).toContain('—');
    expect(footerXml).toContain('w:val="center"');
  });

  test('Chinese official watermark authors VML header part and clears cleanly', async () => {
    const { editor } = openEditor(docx('<w:p><w:r><w:t>内部公文测试</w:t></w:r></w:p>'));

    expect(
      editor.exec({
        type: 'setWatermark',
        watermark: {
          text: '内部资料',
          fontFamily: '宋体',
          color: 'C0C0C0',
          transparency: 0.5,
          layout: 'diagonal',
        },
      }).ok
    ).toBe(true);

    const current = editor.query({ type: 'watermark' });
    expect(current?.text).toBe('内部资料');
    expect(current?.layout).toBe('diagonal');

    const savedWithWm = unzipSync(new Uint8Array(await editor.save()));
    const headerEntries = Object.keys(savedWithWm).filter((k) => /^word\/header\d+\.xml$/.test(k));
    expect(headerEntries.length).toBeGreaterThan(0);
    const combinedHeaders = headerEntries.map((k) => strFromU8(savedWithWm[k]!)).join('\n');
    expect(combinedHeaders).toContain('内部资料');

    expect(editor.exec({ type: 'setWatermark', watermark: null }).ok).toBe(true);
    expect(editor.query({ type: 'watermark' })).toBeNull();
  });

  test('TOC, bookmarks, footnote, endnote, textBox, and clearFormatting work end-to-end', async () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一章 总则</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>第一节 适用范围</w:t></w:r></w:p>' +
      '<w:p><w:r><w:rPr><w:b/><w:color w:val="DC2626"/></w:rPr><w:t>加粗红字正文</w:t></w:r></w:p>';
    const { editor, surface } = openEditor(docx(body, { 'word/styles.xml': HEADING_STYLES_XML }));

    // 1. TOC insert, refresh, remove
    selectPara(surface, 0, 0, 0);
    expect(editor.exec({ type: 'insertToc' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).toContain('TOC');
    expect(editor.exec({ type: 'refreshToc', mode: 'entire' }).ok).toBe(true);
    expect(editor.exec({ type: 'removeToc' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).not.toContain('TOC \\o');

    // 2. Bookmarks
    selectPara(surface, 1, 0, 0);
    expect(editor.exec({ type: 'insertBookmark', name: '第一条锚点' }).ok).toBe(true);
    expect(editor.query({ type: 'bookmarks' }).some((b) => b.name === '第一条锚点')).toBe(true);
    expect(editor.exec({ type: 'goToBookmark', name: '第一条锚点' }).ok).toBe(true);

    // 3. Clear formatting on bold+red paragraph
    selectPara(surface, 2, 0, 6);
    expect(editor.exec({ type: 'clearFormatting' }).ok).toBe(true);
    const afterClear = serializeOoxmlPart(surface.session.part());
    expect(afterClear).not.toContain('w:val="DC2626"');

    // 4. Footnote, endnote, textBox (exitNote returns to body scope between notes)
    selectPara(surface, 2, 6, 6);
    expect(editor.exec({ type: 'insertNote', noteKind: 'footnote' }).ok).toBe(true);
    surface.exitNote?.();
    expect(editor.exec({ type: 'insertNote', noteKind: 'endnote' }).ok).toBe(true);
    surface.exitNote?.();
    expect(editor.exec({ type: 'insertTextBox' }).ok).toBe(true);

    const saved = unzipSync(new Uint8Array(await editor.save()));
    expect(saved['word/footnotes.xml']).toBeDefined();
    expect(saved['word/endnotes.xml']).toBeDefined();
    expect(strFromU8(saved['word/document.xml']!)).toContain('txbxContent');
    expect(strFromU8(saved['word/document.xml']!)).toContain('w:bookmarkStart');

    expect(editor.exec({ type: 'deleteBookmark', name: '第一条锚点' }).ok).toBe(true);
    expect(editor.query({ type: 'bookmarks' }).some((b) => b.name === '第一条锚点')).toBe(false);
  });

  test('spaced table with direct w:tblBorders clears outer box on tableNone/threeLine, keeps tblCellSpacing, and undoes in one step', () => {
    const spacedTbl =
      `<w:tbl><w:tblPr>` +
      `<w:tblCellSpacing w:w="80" w:type="dxa"/>` +
      `<w:tblBorders><w:top w:val="single" w:sz="24" w:color="C01212"/><w:left w:val="single" w:sz="24" w:color="C01212"/><w:bottom w:val="single" w:sz="24" w:color="C01212"/><w:right w:val="single" w:sz="24" w:color="C01212"/></w:tblBorders>` +
      `</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>甲1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙1</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>甲2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const { editor, surface } = openEditor(docx(spacedTbl));
    selectPara(surface, 0, 0, 0);
    const beforeXml = serializeOoxmlPart(surface.session.part());

    expect(editor.exec({ type: 'setTableBorders', scope: 'tableNone' }).ok).toBe(true);
    const afterNone = serializeOoxmlPart(surface.session.part());
    expect(afterNone).toContain('w:tblCellSpacing');
    expect(afterNone).toContain('w:w="80"');
    expect(afterNone).not.toContain('C01212');
    expect(afterNone).not.toContain('w:val="single"');
    const tblBordersMatch = afterNone.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/);
    expect(tblBordersMatch).not.toBeNull();
    expect(hasBorder(tblBordersMatch![0]!, 'top', 'none')).toBe(true);
    expect(hasBorder(tblBordersMatch![0]!, 'bottom', 'none')).toBe(true);
    expect(hasBorder(tblBordersMatch![0]!, 'left', 'none')).toBe(true);
    expect(hasBorder(tblBordersMatch![0]!, 'right', 'none')).toBe(true);

    // Single undo restores red w:tblBorders and original cell borders bit-for-bit
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).toBe(beforeXml);

    // threeLine also clears the red outer w:tblBorders box while keeping tblCellSpacing
    expect(editor.exec({ type: 'setTableBorders', scope: 'threeLine' }).ok).toBe(true);
    const afterThree = serializeOoxmlPart(surface.session.part());
    expect(afterThree).toContain('w:tblCellSpacing');
    expect(afterThree).not.toContain('C01212');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).toBe(beforeXml);
  });

  test('multi-section sectionFurniture query reads caret section and setHeaderFooterOptions targets section vs document accurately', async () => {
    const twoSectionBody =
      `<w:p><w:r><w:t>第一节正文</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:sectPr>` +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="650" w:footer="600"/>` +
      `</w:sectPr></w:pPr></w:p>` +
      `<w:p><w:r><w:t>第二节正文</w:t></w:r></w:p>`;
    const twoSectionDoc = zipSync({
      '[Content_Types].xml': strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${CT}">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
          `</Types>`
      ),
      '_rels/.rels': strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL}">` +
          `<Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/>` +
          `</Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${REL}"/>`
      ),
      'word/document.xml': strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body>${twoSectionBody}` +
          `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="750" w:footer="700"/></w:sectPr>` +
          `</w:body></w:document>`
      ),
    });

    const { editor, surface } = openEditor(twoSectionDoc);
    // Caret in Section 2 (paragraph index 2)
    selectPara(surface, 2, 0, 0);
    expect(editor.query({ type: 'sectionFurniture' })).toEqual({
      sectionIndex: 1,
      sectionCount: 2,
      titlePage: false,
      evenAndOddHeaders: false,
      headerDistanceTwips: 750,
      footerDistanceTwips: 700,
    });

    // Toggling titlePage from section 2 affects ONLY section 2
    expect(editor.exec({ type: 'setHeaderFooterOptions', titlePage: true }).ok).toBe(true);
    expect(editor.query({ type: 'sectionFurniture' })?.titlePage).toBe(true);

    // Section-scoped header distance affects ONLY section 2
    expect(
      editor.exec({ type: 'setHeaderFooterOptions', scope: 'section', headerDistanceTwips: 850 }).ok
    ).toBe(true);
    const savedSection = strFromU8(
      unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!
    );
    const sectPrs = [...savedSection.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(
      (m) => m[0]!
    );
    expect(sectPrs.length).toBe(2);
    expect(sectPrs[0]).not.toContain('<w:titlePg');
    expect(sectPrs[0]).toContain('w:header="650"');
    expect(sectPrs[0]).toContain('w:footer="600"');
    expect(sectPrs[1]).toContain('<w:titlePg');
    expect(sectPrs[1]).toContain('w:header="850"');
    expect(sectPrs[1]).toContain('w:footer="700"');

    // Document-scoped footer distance updates BOTH sections in one undoable step
    expect(
      editor.exec({ type: 'setHeaderFooterOptions', scope: 'document', footerDistanceTwips: 992 })
        .ok
    ).toBe(true);
    const savedDocScope = strFromU8(
      unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!
    );
    const sectPrsAll = [...savedDocScope.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(
      (m) => m[0]!
    );
    expect(sectPrsAll[0]).toContain('w:footer="992"');
    expect(sectPrsAll[1]).toContain('w:footer="992"');

    // Single undo restores both sections' original footer distances
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const afterUndo = strFromU8(
      unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!
    );
    const sectPrsUndone = [...afterUndo.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(
      (m) => m[0]!
    );
    expect(sectPrsUndone[0]).toContain('w:footer="600"');
    expect(sectPrsUndone[1]).toContain('w:footer="700"');
  });

  test('refreshToc pageNumbers preserves custom TOC paragraph style and right-aligned tab stop', async () => {
    const customStyles =
      `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="31"><w:name w:val="CustomToc1"/></w:style>` +
      `</w:styles>`;
    const customTocBody =
      `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>` +
      `<w:p><w:pPr><w:pStyle w:val="31"/><w:tabs><w:tab w:val="right" w:pos="6100"/></w:tabs></w:pPr>` +
      `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:hyperlink w:anchor="_Toc1"><w:r><w:t>一、样件正文</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>9</w:t></w:r></w:hyperlink>` +
      `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
      `</w:sdtContent></w:sdt>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc1"/><w:r><w:t>一、样件正文</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const { editor } = openEditor(docx(customTocBody, { 'word/styles.xml': customStyles }));
    expect(editor.exec({ type: 'refreshToc', mode: 'pageNumbers' }).ok).toBe(true);
    const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
    expect(xml).toContain('<w:pStyle w:val="31"/>');
    expect(xml).toContain('w:pos="6100"');
    expect(xml).toContain('<w:t>1</w:t>');
    expect(xml).not.toContain('<w:t>9</w:t>');

    // Full rebuild (mode: 'entire') ALSO preserves custom style 31 and 6100 twips tab stop
    expect(editor.exec({ type: 'refreshToc', mode: 'entire' }).ok).toBe(true);
    const xmlEntire = strFromU8(
      unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!
    );
    expect(xmlEntire).toContain('<w:pStyle w:val="31"/>');
    expect(xmlEntire).toContain('w:pos="6100"');
    expect(xmlEntire).not.toContain('w:val="TOC1"');
  });

  test('refreshToc entire on mixed-level custom TOC matches existing targets and level templates when a new heading is inserted', async () => {
    const mixedStyles =
      `<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="31"><w:name w:val="CustomToc1"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="32"><w:name w:val="CustomToc2"/></w:style>` +
      `</w:styles>`;
    const mixedBody =
      `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>` +
      `<w:p><w:pPr><w:pStyle w:val="31"/><w:tabs><w:tab w:val="right" w:pos="6100"/></w:tabs></w:pPr>` +
      `<w:r><w:fldChar w:fldCharType="begin"/><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:hyperlink w:anchor="_Toc1600000001"><w:r><w:t>一、样件正文</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r></w:hyperlink></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="32"/><w:tabs><w:tab w:val="right" w:pos="5500"/></w:tabs></w:pPr>` +
      `<w:hyperlink w:anchor="_Toc1600000002"><w:r><w:t>二、样件第二页</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>2</w:t></w:r></w:hyperlink></w:p>` +
      `<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
      `</w:sdtContent></w:sdt>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="2" w:name="_Toc1600000001"/><w:r><w:t>一、样件正文</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>新增一级标题</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:bookmarkStart w:id="3" w:name="_Toc1600000002"/><w:r><w:t>二、样件第二页</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>`;
    const { editor } = openEditor(docx(mixedBody, { 'word/styles.xml': mixedStyles }));
    expect(editor.exec({ type: 'refreshToc', mode: 'entire' }).ok).toBe(true);
    const xml = strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
    const sdtXml = xml.match(/<w:sdtContent>[\s\S]*?<\/w:sdtContent>/)?.[0] ?? '';
    const tocParas = [...sdtXml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)]
      .map((m) => m[0]!)
      .filter((p) => p.includes('<w:hyperlink'));
    expect(tocParas.length).toBe(3);
    // Entry 1 (existing Heading1) -> style 31, tab 6100
    expect(tocParas[0]).toContain('一、样件正文');
    expect(tocParas[0]).toContain('<w:pStyle w:val="31"/>');
    expect(tocParas[0]).toContain('w:pos="6100"');
    // Entry 2 (new Heading1 inserted before Heading2) -> inherits level-1 custom style 31, tab 6100
    expect(tocParas[1]).toContain('新增一级标题');
    expect(tocParas[1]).toContain('<w:pStyle w:val="31"/>');
    expect(tocParas[1]).toContain('w:pos="6100"');
    // Entry 3 (existing Heading2 _Toc1600000002) -> keeps custom style 32, tab 5500
    expect(tocParas[2]).toContain('二、样件第二页');
    expect(tocParas[2]).toContain('<w:pStyle w:val="32"/>');
    expect(tocParas[2]).toContain('w:pos="5500"');
    expect(sdtXml).not.toContain('w:val="TOC1"');
    expect(sdtXml).not.toContain('w:val="TOC2"');
  });
});
