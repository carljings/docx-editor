import { describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '@docx-editor.dev/core/store';
import { buildStyleCascadeTable, resolveParagraphLayoutInputs } from '../style-cascade.ts';
import { paragraphIndent, resolveParagraphFirstLineIndent } from '../paragraph-indent.ts';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function stylesTable(stylesInner: string) {
  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  ${stylesInner}
</w:styles>`;
  const parsed = readOoxmlPart(xml, { name: '/word/styles.xml', contentType: 'application/xml' });
  if (!parsed.ok) throw new Error('failed to parse styles');
  return buildStyleCascadeTable(parsed.part.root);
}

function firstParagraphOf(bodyInner: string) {
  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="${W_NS}"><w:body>${bodyInner}</w:body></w:document>`;
  const parsed = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'application/xml' });
  if (!parsed.ok) throw new Error('failed to parse doc');
  const body = parsed.part.root.children.find((c) => 'localName' in c && c.localName === 'body');
  if (!body || !('children' in body)) throw new Error('missing body');
  const p = body.children.find((c) => 'localName' in c && c.localName === 'p');
  if (!p || !('localName' in p)) throw new Error('missing p');
  return p;
}

describe('w:ind character-unit indentation (§17.3.1.12)', () => {
  test('single attribute: w:firstLineChars alone scales with paragraph font size', () => {
    const p = firstParagraphOf(
      '<w:p><w:pPr><w:ind w:firstLineChars="200"/></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>公文正文</w:t></w:r></w:p>'
    );
    const inputs = resolveParagraphLayoutInputs(p, 410, undefined);
    expect(inputs.indent.firstLine).toBeCloseTo(32, 4);
    expect(inputs.indent.hanging).toBe(0);
  });

  test('dual attribute: non-zero w:firstLineChars supersedes stale w:firstLine twips', () => {
    const p = firstParagraphOf(
      '<w:p><w:pPr><w:ind w:firstLine="200" w:firstLineChars="200"/></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>正文</w:t></w:r></w:p>'
    );
    const inputs = resolveParagraphLayoutInputs(p, 410, undefined);
    expect(inputs.indent.firstLine).toBeCloseTo(32, 4);
  });

  test('zero firstLineChars alongside explicit non-zero twips keeps explicit twips', () => {
    const p = firstParagraphOf(
      '<w:p><w:pPr><w:ind w:left="0" w:leftChars="0" w:firstLine="640" w:firstLineChars="0"/></w:pPr>' +
        '<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>巡查记录</w:t></w:r></w:p>'
    );
    const inputs = resolveParagraphLayoutInputs(p, 410, undefined);
    expect(inputs.indent.firstLine).toBeCloseTo(32, 4);
    expect(inputs.indent.left).toBe(0);
  });

  test('style inheritance chain (Style 2 -> Style 3) resolves 2-char indent and allows zero override', () => {
    const cascade = stylesTable(`
      <w:style w:type="paragraph" w:styleId="3">
        <w:name w:val="Body Text Indent"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:ind w:firstLine="200" w:firstLineChars="200"/></w:pPr>
        <w:rPr><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
      </w:style>
      <w:style w:type="paragraph" w:styleId="2">
        <w:name w:val="Official Body"/>
        <w:basedOn w:val="3"/>
      </w:style>
    `);

    const inheritedP = firstParagraphOf(
      '<w:p><w:pPr><w:pStyle w:val="2"/></w:pPr><w:r><w:t>本次验收核对中文断行</w:t></w:r></w:p>'
    );
    expect(resolveParagraphLayoutInputs(inheritedP, 410, cascade).indent.firstLine).toBeCloseTo(
      32,
      4
    );

    const emptyInheritedP = firstParagraphOf('<w:p><w:pPr><w:pStyle w:val="2"/></w:pPr></w:p>');
    expect(
      resolveParagraphLayoutInputs(emptyInheritedP, 410, cascade).indent.firstLine
    ).toBeCloseTo(32, 4);

    const cancelledCharsOnly = firstParagraphOf(
      '<w:p><w:pPr><w:pStyle w:val="2"/><w:ind w:firstLineChars="0"/></w:pPr><w:r><w:t>无缩进</w:t></w:r></w:p>'
    );
    expect(resolveParagraphLayoutInputs(cancelledCharsOnly, 410, cascade).indent.firstLine).toBe(0);

    const cancelledBoth = firstParagraphOf(
      '<w:p><w:pPr><w:pStyle w:val="2"/><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr><w:r><w:t>无缩进</w:t></w:r></w:p>'
    );
    expect(resolveParagraphLayoutInputs(cancelledBoth, 410, cascade).indent.firstLine).toBe(0);
  });

  test('font size change (16pt -> 22pt) scales firstLineChars from 32pt to 44pt', () => {
    const cascade = stylesTable(`
      <w:style w:type="paragraph" w:styleId="Body2">
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:ind w:firstLine="200" w:firstLineChars="200"/></w:pPr>
        <w:rPr><w:sz w:val="32"/></w:rPr>
      </w:style>
    `);
    const p16 = firstParagraphOf(
      '<w:p><w:pPr><w:pStyle w:val="Body2"/></w:pPr><w:r><w:t>三号正文</w:t></w:r></w:p>'
    );
    const p22 = firstParagraphOf(
      '<w:p><w:pPr><w:pStyle w:val="Body2"/></w:pPr><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>二号标题</w:t></w:r></w:p>'
    );
    expect(resolveParagraphLayoutInputs(p16, 410, cascade).indent.firstLine).toBeCloseTo(32, 4);
    expect(resolveParagraphLayoutInputs(p22, 410, cascade).indent.firstLine).toBeCloseTo(44, 4);
  });

  test('leftChars, rightChars, hangingChars and trust-boundary clamp behave safely', () => {
    const side = paragraphIndent(
      [{ localName: 'ind', attributes: { left: '100', leftChars: '200', rightChars: '100' } }],
      10.5
    );
    expect(side.left).toBeCloseTo(21, 4);
    expect(side.right).toBeCloseTo(10.5, 4);

    const hang = resolveParagraphFirstLineIndent(
      [{ localName: 'ind', attributes: { hanging: '19', hangingChars: '150' } }],
      16
    );
    expect(hang.hanging).toBeCloseTo(24, 4);
    expect(hang.firstLine).toBe(0);

    const clamped = resolveParagraphFirstLineIndent(
      [{ localName: 'ind', attributes: { firstLineChars: '999999999' } }],
      16
    );
    expect(clamped.firstLine).toBe(1584);

    const rejected = resolveParagraphFirstLineIndent(
      [{ localName: 'ind', attributes: { firstLineChars: '1234567890' } }],
      16
    );
    expect(rejected.firstLine).toBe(0);
  });
});
