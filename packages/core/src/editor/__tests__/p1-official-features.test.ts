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

function docx(bodyInner: string): Uint8Array {
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
        `<w:sectPr><w:pgSz w:w="12500" w:h="17000"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>` +
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
    expect(xml).toContain('<w:cols w:num="2" w:sep="1" w:space="425"/>');
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
