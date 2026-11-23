// P1 WPS official-document editing capabilities in Core:
// Chinese numbering presets, paragraph borders & shading, underline variants,
// character spacing, and multi-column section page setup.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { formatNumFmt } from '@docx-editor.dev/core/layout';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(bodyInner: string, sectPrInner?: string): Uint8Array {
  const sect =
    sectPrInner ??
    '<w:pgSz w:w="12500" w:h="17000"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>';
  return zipSync({
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
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body>${bodyInner}` +
        `<w:sectPr>${sect}</w:sectPr>` +
        `</w:body></w:document>`
    ),
  });
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

function selectPara(surface: PaginatedSurface, index: number, from: number, to: number): void {
  const id = surface.session.paragraphIds()[index]!;
  surface.setSelection({
    anchor: { paragraphId: id, offset: from },
    head: { paragraphId: id, offset: to },
  });
}

describe('P1 Chinese official numbering', () => {
  test('formatNumFmt formats Chinese counting numerals accurately', () => {
    expect(formatNumFmt('chineseCountingThousand', 1)).toBe('一');
    expect(formatNumFmt('chineseCountingThousand', 2)).toBe('二');
    expect(formatNumFmt('chineseCountingThousand', 10)).toBe('十');
    expect(formatNumFmt('chineseCountingThousand', 11)).toBe('十一');
    expect(formatNumFmt('chineseCountingThousand', 20)).toBe('二十');
    expect(formatNumFmt('chineseCountingThousand', 21)).toBe('二十一');
    expect(formatNumFmt('chineseCountingThousand', 100)).toBe('一百');
    expect(formatNumFmt('chineseCountingThousand', 101)).toBe('一百零一');
    expect(formatNumFmt('chineseLegalSimplified', 10)).toBe('壹拾');
  });

  test('toggleList official presets author numbering.xml and support switching & clearing', async () => {
    const { editor, surface } = openEditor(
      docx(
        '<w:p><w:r><w:t>重点工作部署</w:t></w:r></w:p><w:p><w:r><w:t>督办落实要求</w:t></w:r></w:p>'
      )
    );
    editor.exec({ type: 'selectAll' });
    const applied = editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading1' });
    expect(applied.ok).toBe(true);

    const saved = unzipSync(new Uint8Array(await editor.save()));
    const numberingXml = strFromU8(saved['word/numbering.xml']!);
    expect(numberingXml).toContain('w:val="chineseCountingThousand"');
    expect(numberingXml).toContain('w:val="%1、"');
    expect(serializeOoxmlPart(surface.session.part())).toContain('<w:numPr>');

    const switched = editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading2' });
    expect(switched.ok).toBe(true);
    const numberingAfter = strFromU8(
      unzipSync(new Uint8Array(await editor.save()))['word/numbering.xml']!
    );
    expect(numberingAfter).toContain('w:val="（%1）"');

    const cleared = editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'clear' });
    expect(cleared.ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).not.toContain('<w:numPr>');
  });

  test('reapplying same preset preserves numId/ilvl, joins adjacent list, preserves sublevels on switch, and restarts explicitly', () => {
    const { editor, surface } = openEditor(
      docx(
        '<w:p><w:r><w:t>甲组首项</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>甲组次项</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>甲组末项</w:t></w:r></w:p>'
      )
    );
    const readNumIds = () => {
      const xml = serializeOoxmlPart(surface.session.part());
      return [
        ...xml.matchAll(/<w:numPr><w:ilvl w:val="(\d+)"\/><w:numId w:val="(\d+)"\/><\/w:numPr>/g),
      ].map((m) => ({ ilvl: Number(m[1]), numId: m[2]! }));
    };

    editor.exec({ type: 'selectAll' });
    expect(editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading1' }).ok).toBe(
      true
    );
    const initial = readNumIds();
    expect(initial.length).toBe(3);
    expect(new Set(initial.map((x) => x.numId)).size).toBe(1);
    const origNumId = initial[0]!.numId;

    // 1. Reapplying same preset on middle paragraph preserves numId & level (no-op)
    selectPara(surface, 1, 0, 0);
    const reapply = editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading1' });
    expect(reapply.ok).toBe(true);
    expect(reapply.changed).toBe(false);
    expect(readNumIds()).toEqual([
      { ilvl: 0, numId: origNumId },
      { ilvl: 0, numId: origNumId },
      { ilvl: 0, numId: origNumId },
    ]);

    // 2. Clearing P2 then applying cn-heading1 rejoins adjacent list numId
    expect(editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'clear' }).ok).toBe(true);
    expect(readNumIds().length).toBe(2);
    expect(editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading1' }).ok).toBe(
      true
    );
    expect(readNumIds()).toEqual([
      { ilvl: 0, numId: origNumId },
      { ilvl: 0, numId: origNumId },
      { ilvl: 0, numId: origNumId },
    ]);

    // 3. Demote P2 to sublevel ilvl=1, reapply cn-heading1 across all 3 -> keeps numId & ilvl=1
    selectPara(surface, 1, 0, 0);
    expect(editor.exec({ type: 'adjustIndent', direction: 'increase' }).ok).toBe(true);
    expect(readNumIds()[1]).toEqual({ ilvl: 1, numId: origNumId });
    editor.exec({ type: 'selectAll' });
    const reapplySub = editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-heading1' });
    expect(reapplySub.ok).toBe(true);
    expect(reapplySub.changed).toBe(false);
    expect(readNumIds()).toEqual([
      { ilvl: 0, numId: origNumId },
      { ilvl: 1, numId: origNumId },
      { ilvl: 0, numId: origNumId },
    ]);

    // 4. Switch preset to cn-decimal across all 3 -> single new numId preserving ilvl=[0, 1, 0]
    expect(editor.exec({ type: 'toggleList', kind: 'ordered', preset: 'cn-decimal' }).ok).toBe(
      true
    );
    const switched = readNumIds();
    expect(switched.length).toBe(3);
    expect(switched[0]!.numId).not.toBe(origNumId);
    expect(new Set(switched.map((x) => x.numId)).size).toBe(1);
    expect(switched.map((x) => x.ilvl)).toEqual([0, 1, 0]);

    // 5. Explicit restart on P2 mints a fresh numId at ilvl=0
    selectPara(surface, 1, 0, 0);
    expect(editor.exec({ type: 'toggleList', kind: 'ordered', restart: true }).ok).toBe(true);
    const restarted = readNumIds();
    expect(restarted[1]!.numId).not.toBe(switched[0]!.numId);
    expect(restarted[1]!.ilvl).toBe(0);
  });
});

describe('P1 paragraph borders, shading, underline variants, character spacing, and columns', () => {
  test('setParagraphFormat applies and clears w:pBdr and w:shd atomically', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>红头标题</w:t></w:r></w:p>'));
    selectPara(surface, 0, 0, 4);
    const res = editor.exec({
      type: 'setParagraphFormat',
      borders: [
        { edge: 'bottom', attributes: { val: 'single', sz: '18', space: '4', color: 'DC2626' } },
      ],
      shading: 'FEF2F2',
    });
    expect(res.ok).toBe(true);
    const xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('<w:pBdr>');
    expect(xml).toContain('w:color="DC2626"');
    expect(xml).toContain('<w:shd');
    expect(xml).toContain('w:fill="FEF2F2"');

    const cleared = editor.exec({ type: 'setParagraphFormat', borders: [], shading: null });
    expect(cleared.ok).toBe(true);
    const afterXml = serializeOoxmlPart(surface.session.part());
    expect(afterXml).not.toContain('<w:pBdr>');
    expect(afterXml).not.toContain('<w:shd');
  });

  test('setParagraphFormat borderMode merge updates only target edge and preserves existing box sides', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>既有外框段落</w:t></w:r></w:p>'));
    selectPara(surface, 0, 0, 4);
    const boxEdge = { val: 'single', sz: '6', space: '4', color: '64748B' };
    expect(
      editor.exec({
        type: 'setParagraphFormat',
        borderMode: 'replace',
        borders: [
          { edge: 'top', attributes: boxEdge },
          { edge: 'left', attributes: boxEdge },
          { edge: 'bottom', attributes: boxEdge },
          { edge: 'right', attributes: boxEdge },
        ],
      }).ok
    ).toBe(true);

    expect(
      editor.exec({
        type: 'setParagraphFormat',
        borderMode: 'merge',
        borders: [
          { edge: 'bottom', attributes: { val: 'single', sz: '18', space: '4', color: 'DC2626' } },
        ],
      }).ok
    ).toBe(true);

    const xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('<w:top w:color="64748B" w:space="4" w:sz="6" w:val="single"/>');
    expect(xml).toContain('<w:left w:color="64748B" w:space="4" w:sz="6" w:val="single"/>');
    expect(xml).toContain('<w:bottom w:color="DC2626" w:space="4" w:sz="18" w:val="single"/>');
    expect(xml).toContain('<w:right w:color="64748B" w:space="4" w:sz="6" w:val="single"/>');

    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    const undoneXml = serializeOoxmlPart(surface.session.part());
    expect(undoneXml).toContain('<w:bottom w:color="64748B" w:space="4" w:sz="6" w:val="single"/>');
  });

  test('setMarkAttr writes underline line styles and character spacing', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>测试文字样式</w:t></w:r></w:p>'));
    selectPara(surface, 0, 0, 6);
    expect(
      editor.exec({ type: 'setMarkAttr', mark: 'underline', attr: 'val', value: 'wave' }).ok
    ).toBe(true);
    expect(
      editor.exec({ type: 'setMarkAttr', mark: 'characterSpacing', attr: 'val', value: 40 }).ok
    ).toBe(true);
    const xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('<w:u w:val="wave"/>');
    expect(xml).toContain('<w:spacing w:val="40"/>');
    expect(editor.snapshot().formatting?.underline).toBe(true);

    expect(
      editor.exec({ type: 'setMarkAttr', mark: 'underline', attr: 'val', value: 'none' }).ok
    ).toBe(true);
    expect(editor.snapshot().formatting?.underline).toBe(false);
  });

  test('setPageSetup writes w:cols while preserving custom paper dimensions', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>双栏正文内容</w:t></w:r></w:p>'));
    const res = editor.exec({
      type: 'setPageSetup',
      columns: { count: 2, spaceTwips: 425, separator: true },
      scope: 'section',
    });
    expect(res.ok).toBe(true);
    const xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('w:w="12500"');
    expect(xml).toContain('w:h="17000"');
    expect(xml).toContain('w:num="2"');
    expect(xml).toContain('w:sep="1"');
    expect(xml).toContain('w:space="425"');
    expect(editor.snapshot().pageSetup?.columns?.count).toBe(2);
    expect(editor.snapshot().pageSetup?.columns?.separator).toBe(true);

    expect(
      editor.exec({
        type: 'setPageSetup',
        columns: { count: 1, spaceTwips: 425 },
        scope: 'section',
      }).ok
    ).toBe(true);
    expect(editor.snapshot().pageSetup?.columns?.count).toBe(1);
  });

  test('setPageSetup patches w:sep losslessly on unequal, zero-gap, and 4-column sections', () => {
    const unequal = openEditor(
      docx(
        '<w:p><w:r><w:t>不等宽双栏</w:t></w:r></w:p>',
        '<w:pgSz w:w="12500" w:h="17000"/>' +
          '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>' +
          '<w:cols w:num="2" w:space="0" w:equalWidth="0">' +
          '<w:col w:w="3000" w:space="400"/><w:col w:w="6900"/>' +
          '</w:cols>'
      )
    );
    expect(unequal.editor.snapshot().pageSetup?.columns).toEqual({
      count: 2,
      gapTwips: 0,
      equalWidth: false,
    });
    expect(
      unequal.editor.exec({
        type: 'setPageSetup',
        columns: { separator: true },
        scope: 'section',
      }).ok
    ).toBe(true);
    const unequalXml = serializeOoxmlPart(unequal.surface.session.part());
    expect(unequalXml).toContain('w:equalWidth="0"');
    expect(unequalXml).toContain('w:space="0"');
    expect(unequalXml).toContain('w:sep="1"');
    expect(unequalXml).toContain('<w:col w:space="400" w:w="3000"/>');
    expect(unequalXml).toContain('<w:col w:w="6900"/>');
    expect(unequalXml).toContain('w:w="12500"');

    const fourCols = openEditor(
      docx(
        '<w:p><w:r><w:t>四栏零间距</w:t></w:r></w:p>',
        '<w:pgSz w:w="12500" w:h="17000"/>' +
          '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>' +
          '<w:cols w:num="4" w:space="0"/>'
      )
    );
    expect(fourCols.editor.snapshot().pageSetup?.columns).toEqual({ count: 4, gapTwips: 0 });
    expect(
      fourCols.editor.exec({
        type: 'setPageSetup',
        columns: { separator: true },
        scope: 'section',
      }).ok
    ).toBe(true);
    const fourXml = serializeOoxmlPart(fourCols.surface.session.part());
    expect(fourXml).toContain('w:num="4"');
    expect(fourXml).toContain('w:space="0"');
    expect(fourXml).toContain('w:sep="1"');
  });

  test('toggleHeaderRow works with w:cantSplit, multi-row headers, and live TableContext state', () => {
    const { editor, surface } = openEditor(
      docx(
        '<w:tbl>' +
          '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
          '<w:tr><w:trPr><w:cantSplit/></w:trPr>' +
          '<w:tc><w:p><w:r><w:t>表头1</w:t></w:r></w:p></w:tc>' +
          '<w:tc><w:p><w:r><w:t>列2</w:t></w:r></w:p></w:tc></w:tr>' +
          '<w:tr><w:tc><w:p><w:r><w:t>第二行</w:t></w:r></w:p></w:tc>' +
          '<w:tc><w:p><w:r><w:t>数据2</w:t></w:r></w:p></w:tc></w:tr>' +
          '<w:tr><w:tc><w:p><w:r><w:t>第三行</w:t></w:r></w:p></w:tc>' +
          '<w:tc><w:p><w:r><w:t>数据3</w:t></w:r></w:p></w:tc></w:tr>' +
          '</w:tbl>'
      )
    );
    selectPara(surface, 0, 0, 0);
    expect(editor.snapshot().table?.headerRowCount).toBeUndefined();
    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    let xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('<w:cantSplit/>');
    expect(xml).toContain('<w:tblHeader w:val="1"/>');
    expect(editor.snapshot().table?.headerRowCount).toBe(1);
    expect(editor.snapshot().table?.isHeaderRow).toBe(true);

    // Move caret to row 1 and extend repeating header to [1, 2]
    selectPara(surface, 2, 0, 0);
    expect(editor.snapshot().table?.isHeaderRow).toBe(false);
    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    expect(editor.snapshot().table?.headerRowCount).toBe(2);
    expect(editor.snapshot().table?.isHeaderRow).toBe(true);

    // Move caret back to row 0: off ([1, 2] -> []) then on again ([] -> [1]), preserving cantSplit
    selectPara(surface, 0, 0, 0);
    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    xml = serializeOoxmlPart(surface.session.part());
    expect(xml).not.toContain('tblHeader');
    expect(xml).toContain('<w:cantSplit/>');
    expect(editor.snapshot().table?.headerRowCount).toBeUndefined();

    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    xml = serializeOoxmlPart(surface.session.part());
    expect(xml).toContain('<w:tblHeader w:val="1"/>');
    expect(xml).toContain('<w:cantSplit/>');
    expect(editor.snapshot().table?.headerRowCount).toBe(1);
    expect(editor.snapshot().table?.isHeaderRow).toBe(true);
  });

  test('toggleHeaderRow toggles repeating header row on and off via setTableProperties', () => {
    const { editor, surface } = openEditor(docx('<w:p><w:r><w:t>表前段落</w:t></w:r></w:p>'));
    expect(editor.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true);
    expect(editor.can({ type: 'toggleHeaderRow' }).ok).toBe(true);
    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).toContain('<w:tblHeader w:val="1"/>');

    expect(editor.exec({ type: 'toggleHeaderRow' }).ok).toBe(true);
    expect(serializeOoxmlPart(surface.session.part())).not.toContain('tblHeader');
  });
});
