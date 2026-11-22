// Whole-run selection replacement formatting preservation & WPS China Edition
// clean follower-style paragraph creation on Enter (`w:next`).

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { setWpsCompatMode } from '@docx-editor.dev/core/layout';
import {
  directParagraphProperties,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const STY = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

/** Official-document style hierarchy matching WPS template: Style 1 -> w:next 2 -> basedOn 3. */
const OFFICIAL_STYLES =
  `<w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>` +
  `</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="1">` +
  `<w:name w:val="Normal"/><w:next w:val="2"/>` +
  `<w:pPr><w:widowControl w:val="0"/><w:jc w:val="both"/></w:pPr>` +
  `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体" w:cs="Times New Roman"/>` +
  `<w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="2">` +
  `<w:name w:val="Body Text First Indent 2"/><w:basedOn w:val="3"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="3">` +
  `<w:name w:val="Body Text Indent"/><w:basedOn w:val="1"/>` +
  `<w:pPr><w:ind w:firstLine="200" w:firstLineChars="200"/></w:pPr>` +
  `<w:rPr><w:rFonts w:eastAsia="仿宋_GB2312"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>`;

/** Title/header paragraph carrying direct overrides + a 16pt Heiti run over a 22pt mark. */
const TEMPLATE_P0 =
  `<w:p><w:pPr>` +
  `<w:spacing w:line="560" w:lineRule="exact"/>` +
  `<w:ind w:left="0" w:leftChars="0"/>` +
  `<w:jc w:val="left"/>` +
  `<w:rPr><w:rFonts w:hint="default" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体" w:cs="Times New Roman"/>` +
  `<w:sz w:val="44"/><w:szCs w:val="44"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>` +
  `</w:pPr>` +
  `<w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体" w:cs="Times New Roman"/>` +
  `<w:sz w:val="32"/><w:szCs w:val="32"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>` +
  `<w:t>xxxxx</w:t></w:r></w:p>`;

function docx(body: string, styles = OFFICIAL_STYLES): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId2" Type="${STY}" Target="styles.xml"/></Relationships>`
    ),
    'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}">${styles}</w:styles>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const mounted: { editor: DocxEditorInstance; container: HTMLElement }[] = [];

afterEach(() => {
  setWpsCompatMode(false);
  for (const entry of mounted.splice(0)) {
    entry.editor.destroy();
    entry.container.remove();
  }
});

function mount(bytes: Uint8Array): { editor: DocxEditorInstance; surface: PaginatedSurface } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: bytes });
  mounted.push({ editor, container });
  if (!editor.surface) throw new Error('surface failed to mount');
  return { editor, surface: editor.surface };
}

function select(
  surface: PaginatedSurface,
  pIndex: number,
  start: number,
  end: number,
  endPIndex = pIndex
): void {
  const ids = surface.session.paragraphIds();
  surface.setSelection({
    anchor: { paragraphId: ids[pIndex]!, offset: start },
    head: { paragraphId: ids[endPIndex]!, offset: end },
  });
}

function bodyParagraphXmls(surface: PaginatedSurface): string[] {
  const xml = serializeOoxmlPart(surface.session.part());
  return xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>|<w:p\/>/g) ?? [];
}

describe('whole-run and range selection replacement preserves run properties', () => {
  test('typing over an entire styled run retains w:sz, w:rFonts, and w:lang intact', async () => {
    const { editor, surface } = mount(docx(TEMPLATE_P0));
    select(surface, 0, 0, 5);
    surface.type('ABCDE');

    const p0Id = surface.session.paragraphIds()[0]!;
    expect(paragraphTextOf(surface.session.part(), p0Id)).toBe('ABCDE');
    expect(surface.formatting().fontSizeHalfPoints).toBe(32);

    const p0Xml = bodyParagraphXmls(surface)[0]!;
    expect(p0Xml).toContain('<w:sz w:val="32"/>');
    expect(p0Xml).toContain('<w:szCs w:val="32"/>');
    expect(p0Xml).toContain('w:eastAsia="黑体"');
    expect(p0Xml).toContain('w:eastAsia="zh-CN"');
    expect(p0Xml).toContain('<w:t>ABCDE</w:t>');

    const saved = unzipSync(new Uint8Array(await editor.save()));
    const savedXml = new TextDecoder().decode(saved['word/document.xml']!);
    expect(savedXml).toMatch(
      /<w:r><w:rPr>[\s\S]*<w:sz w:val="32"\/>[\s\S]*<\/w:rPr><w:t>ABCDE<\/w:t><\/w:r>/
    );
  });

  test('insertPlainText over an entire styled run retains its run formatting (single and multi-line)', () => {
    const { surface } = mount(docx(TEMPLATE_P0));
    select(surface, 0, 0, 5);
    surface.insertPlainText('第一行\n第二行');

    const xmls = bodyParagraphXmls(surface);
    expect(xmls).toHaveLength(2);
    expect(xmls[0]).toContain('<w:sz w:val="32"/>');
    expect(xmls[0]).toContain('第一行');
    expect(xmls[1]).toContain('<w:sz w:val="32"/>');
    expect(xmls[1]).toContain('第二行');
  });

  test('replacing the second run in a multi-run paragraph keeps the replaced run face, not the preceding run', () => {
    const body =
      `<w:p>` +
      `<w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>Lead </w:t></w:r>` +
      `<w:r><w:rPr><w:i/><w:sz w:val="28"/></w:rPr><w:t>Tail</w:t></w:r>` +
      `</w:p>`;
    const { surface } = mount(docx(body));
    select(surface, 0, 5, 9);
    surface.type('Next');

    expect(surface.formatting().fontSizeHalfPoints).toBe(28);
    expect(surface.formatting().italic).toBe(true);
    expect(surface.formatting().bold).toBe(false);
  });

  test('replacing across two paragraphs preserves the first paragraph replaced run formatting', () => {
    const body = `${TEMPLATE_P0}<w:p><w:r><w:t>Second</w:t></w:r></w:p>`;
    const { surface } = mount(docx(body));
    select(surface, 0, 0, 3, 1);
    surface.type('ABC');

    expect(surface.session.paragraphIds()).toHaveLength(1);
    expect(paragraphTextOf(surface.session.part(), surface.session.paragraphIds()[0]!)).toBe(
      'ABCond'
    );
    const p0Xml = bodyParagraphXmls(surface)[0]!;
    expect(p0Xml).toMatch(
      /<w:r><w:rPr>[\s\S]*<w:sz w:val="32"\/>[\s\S]*<\/w:rPr><w:t>ABC<\/w:t><\/w:r>/
    );
  });
});

describe('WPS compat Enter into a different w:next follower style', () => {
  test('resets direct paragraph overrides and run formatting when switching to Style 2', async () => {
    setWpsCompatMode(true);
    const { editor, surface } = mount(docx(TEMPLATE_P0));

    // Step 1: replace whole run "xxxxx" with "ABCDE"
    select(surface, 0, 0, 5);
    surface.type('ABCDE');

    // Step 2: press Enter at end of P0 (Style 1 -> w:next Style 2)
    surface.splitParagraph();
    expect(surface.session.paragraphIds()).toHaveLength(2);

    const p1Id = surface.session.paragraphIds()[1]!;
    const p1Direct = directParagraphProperties(surface.session.part(), p1Id).map(
      (p) => `${p.localName}=${p.attributes?.val ?? ''}`
    );
    expect(p1Direct).toEqual(['pStyle=2']);

    // Caret in the new paragraph immediately inherits Style 2 -> Style 3 (16pt 仿宋, 2-char indent)
    expect(surface.formatting().fontSizeHalfPoints).toBe(32);

    // Step 3: type body text + metric placeholder into P1
    surface.type('本次验收核对中文断行。');
    surface.type('【指标：接报警情总数】');

    const xmls = bodyParagraphXmls(surface);
    expect(xmls[0]).toContain('<w:sz w:val="32"/>');
    expect(xmls[0]).toContain('<w:t>ABCDE</w:t>');
    // P1 must have clean Style 2 pPr and clean run (no direct ind/spacing/jc/Heiti/sz=44)
    expect(xmls[1]).toContain('<w:pPr><w:pStyle w:val="2"/></w:pPr>');
    expect(xmls[1]).not.toContain('w:ind');
    expect(xmls[1]).not.toContain('w:spacing');
    expect(xmls[1]).not.toContain('w:jc');
    expect(xmls[1]).not.toContain('黑体');
    expect(xmls[1]).not.toContain('w:sz w:val="44"');
    expect(xmls[1]).toContain('<w:r><w:t>本次验收核对中文断行。【指标：接报警情总数】</w:t></w:r>');

    // Survives save & reopen
    const savedBytes = new Uint8Array(await editor.save());
    const reopened = mount(savedBytes);
    expect(reopened.surface.session.paragraphIds()).toHaveLength(2);
    expect(
      directParagraphProperties(
        reopened.surface.session.part(),
        reopened.surface.session.paragraphIds()[1]!
      ).map((p) => p.localName)
    ).toEqual(['pStyle']);
  });

  test('preserves trailing w:sectPr when resetting follower-style paragraph properties', () => {
    setWpsCompatMode(true);
    const body =
      `<w:p><w:pPr>` +
      `<w:spacing w:line="560" w:lineRule="exact"/>` +
      `<w:ind w:left="0"/><w:jc w:val="left"/>` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>` +
      `</w:pPr><w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>Title</w:t></w:r></w:p>`;
    const { surface } = mount(docx(body));
    select(surface, 0, 5, 5);
    surface.splitParagraph();

    const xmls = bodyParagraphXmls(surface);
    expect(xmls[0]).not.toContain('w:sectPr');
    expect(xmls[1]).toContain('<w:pStyle w:val="2"/>');
    expect(xmls[1]).toContain('<w:sectPr>');
    expect(xmls[1]).not.toContain('w:ind');
    expect(xmls[1]).not.toContain('w:spacing');
  });

  test('keeps direct formatting when pressing Enter in a style that does not change via w:next', () => {
    setWpsCompatMode(true);
    const body =
      `<w:p><w:pPr><w:pStyle w:val="2"/><w:jc w:val="center"/></w:pPr>` +
      `<w:r><w:rPr><w:b/><w:sz w:val="36"/></w:rPr><w:t>Body</w:t></w:r></w:p>`;
    const { surface } = mount(docx(body));
    select(surface, 0, 4, 4);
    surface.splitParagraph();
    surface.type('Next');

    const xmls = bodyParagraphXmls(surface);
    expect(xmls[1]).toContain('<w:pStyle w:val="2"/>');
    expect(xmls[1]).toContain('<w:jc w:val="center"/>');
    expect(surface.formatting().bold).toBe(true);
    expect(surface.formatting().fontSizeHalfPoints).toBe(36);
  });
});

describe('character-unit first-line indent command & read-only copyFlavours query', () => {
  test('setParagraphFormat indentFirstLineChars scales dynamically on 16pt -> 22pt font size change', () => {
    const body = `<w:p><w:r><w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr><w:t>公文首行缩进测试</w:t></w:r></w:p>`;
    const { editor, surface } = mount(docx(body));
    select(surface, 0, 0, 8);
    const applied = editor.exec({
      type: 'setParagraphFormat',
      indentFirstLineChars: 200,
      indentFirstLineTwips: 640,
    });
    expect(applied.ok).toBe(true);
    expect(editor.snapshot().formatting?.indent?.firstLine).toBe(640);
    expect(bodyParagraphXmls(surface)[0]).toContain('w:firstLineChars="200"');

    // Enlarge text from 16pt -> 22pt: 2 chars must widen from 640 twips to 880 twips.
    select(surface, 0, 0, 8);
    const resized = editor.exec({ type: 'setMarkAttr', mark: 'fontSize', attr: 'val', value: 44 });
    expect(resized.ok).toBe(true);
    expect(editor.snapshot().formatting?.indent?.firstLine).toBe(880);

    // Zeroing via indentFirstLineChars: 0 cancels both char and twip attributes.
    editor.exec({ type: 'setParagraphFormat', indentFirstLineChars: 0, indentFirstLineTwips: 0 });
    expect(editor.snapshot().formatting?.indent?.firstLine).toBe(0);
  });

  test('copyFlavours query returns plain and rich HTML without mutating the document', () => {
    const body = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗标题</w:t></w:r></w:p>`;
    const { editor, surface } = mount(docx(body));
    select(surface, 0, 0, 4);
    const flavours = editor.query({ type: 'copyFlavours' });
    expect(flavours).not.toBeNull();
    expect(flavours?.text).toBe('加粗标题');
    expect(flavours?.html).toContain('加粗标题');
    expect(paragraphTextOf(surface.session.part(), surface.session.paragraphIds()[0]!)).toBe(
      '加粗标题'
    );
  });
});
