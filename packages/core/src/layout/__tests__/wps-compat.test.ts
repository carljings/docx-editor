import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { buildStyleCascadeTable } from '../style-cascade.ts';
import { resolveCjkTypography } from '../cjk-typography.ts';
import { NO_THEME_FONTS } from '../run-style.ts';
import { layoutSemanticDocument } from '../semantic-layout.ts';
import { linesOf, type TextMeasurer } from '../semantic-records.ts';
import { getWpsCompatMode, setWpsCompatMode } from '../wps-compat.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STYLES = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia="宋体"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
const SETTINGS =
  `<w:settings xmlns:w="${W}">` +
  `<w:characterSpacingControl w:val="compressPunctuation"/>` +
  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>` +
  `</w:settings>`;

const INK: Record<string, [number, number]> = {
  '。': [0.14, 0.36],
  '，': [0.14, 0.3],
  '、': [0.14, 0.33],
  '（': [0.64, 0.9],
  '）': [0.1, 0.36],
};

function part(xml: string, name = '/word/document.xml') {
  const result = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const measurer: TextMeasurer = {
  measure: (text, style) =>
    [...text].reduce((sum, ch) => sum + (ch === ' ' ? style.fontSizePt / 2 : style.fontSizePt), 0) +
    text.length * style.characterSpacingPt,
  lineMetrics: (style) => ({
    height: style.fontSizePt * 1.2,
    baseline: style.fontSizePt * 0.85,
  }),
  inkBounds: (text, style) => {
    const pair = INK[text];
    return pair
      ? { left: pair[0] * style.fontSizePt, right: pair[1] * style.fontSizePt }
      : { left: 0.05 * style.fontSizePt, right: 0.95 * style.fontSizePt };
  },
};

afterEach(() => {
  setWpsCompatMode(false);
});

describe('WPS China Edition layout compatibility mode (opt-in)', () => {
  test('defaults to disabled so MS Word parity is unchanged', () => {
    expect(getWpsCompatMode()).toEqual({ leftAlignedPunctuationCompression: false });
    const styles = part(STYLES, '/word/styles.xml').root;
    const settings = part(SETTINGS, '/word/settings.xml').root;
    const cascade = buildStyleCascadeTable(styles, NO_THEME_FONTS, settings);
    expect(cascade.typography?.wpsLeftPunctuationCompression).toBeUndefined();
    const typo = resolveCjkTypography(
      [{ namespaceUri: W, localName: 'jc', attributes: { val: 'left' } }],
      cascade.typography
    );
    expect(typo.hangingPunctuation).toBe(false);
  });

  test('enables wpsLeftPunctuationCompression and invalidates style cascade cacheToken when toggled', () => {
    const styles = part(STYLES, '/word/styles.xml').root;
    const settings = part(SETTINGS, '/word/settings.xml').root;
    const baseline = buildStyleCascadeTable(styles, NO_THEME_FONTS, settings);

    setWpsCompatMode(true);
    expect(getWpsCompatMode()).toEqual({ leftAlignedPunctuationCompression: true });
    const wpsCascade = buildStyleCascadeTable(styles, NO_THEME_FONTS, settings);

    expect(wpsCascade.typography?.wpsLeftPunctuationCompression).toBe(true);
    expect(wpsCascade.cacheToken).not.toBe(baseline.cacheToken);

    const typo = resolveCjkTypography(
      [{ namespaceUri: W, localName: 'jc', attributes: { val: 'left' } }],
      wpsCascade.typography
    );
    expect(typo.hangingPunctuation).toBe(true);
  });

  test('compresses left-aligned Chinese punctuation to fit 28 full-width chars in 442.2pt when enabled', () => {
    // Mirrors P[007] line 2 of 日盯办报告-全指标模板.docx (28 CJK chars including 6 punctuation marks in 442.2pt)
    const sample = '商盯办前一日执法工作情况，法制、情指、经侦、治安、刑侦、';
    const docXml =
      `<w:document xmlns:w="${W}"><w:body>` +
      `<w:p><w:pPr><w:jc w:val="left"/><w:overflowPunct w:val="0"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr><w:t>${sample}</w:t></w:r>` +
      `</w:p>` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2041" w:right="1531" w:bottom="2041" w:left="1531"/></w:sectPr>` +
      `</w:body></w:document>`;

    const doc = part(docXml);
    const styles = part(STYLES, '/word/styles.xml').root;
    const settings = part(SETTINGS, '/word/settings.xml').root;
    const theme = { ...NO_THEME_FONTS, eastAsianFontTableFaces: ['宋体'] };

    // Default MS Word mode: wraps onto 2 lines (27 chars + 1 char)
    setWpsCompatMode(false);
    const wordLayout = layoutSemanticDocument(doc, 0, {
      measurer,
      styleCascade: buildStyleCascadeTable(styles, theme, settings),
    });
    expect(linesOf(wordLayout).length).toBe(2);

    // WPS compat mode: compresses CJK punctuation so all 28 chars fit on 1 line
    setWpsCompatMode(true);
    const wpsLayout = layoutSemanticDocument(doc, 0, {
      measurer,
      styleCascade: buildStyleCascadeTable(styles, theme, settings),
    });
    expect(linesOf(wpsLayout).length).toBe(1);
  });

  test('allows mixed ASCII+CJK optical compression on left-aligned paragraphs even with w:noPunctuationKerning=1', () => {
    const settingsWithNoKern =
      `<w:settings xmlns:w="${W}">` +
      `<w:noPunctuationKerning w:val="1"/>` +
      `<w:characterSpacingControl w:val="compressPunctuation"/>` +
      `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>` +
      `</w:settings>`;
    const docXml =
      `<w:document xmlns:w="${W}"><w:body>` +
      `<w:p><w:pPr><w:jc w:val="left"/><w:overflowPunct w:val="0"/><w:ind w:firstLine="640"/></w:pPr>` +
      `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="32"/></w:rPr>` +
      `<w:t>{{报告日期-1}}，经对采取刑事强制措施的在办案件（{{在办（采取强制措施）案件数}}件）进行全面检查，未发现问题。</w:t>` +
      `</w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2041" w:right="1531" w:bottom="2041" w:left="1531"/></w:sectPr>` +
      `</w:body></w:document>`;

    const measureMixed = (text: string, style: Parameters<TextMeasurer['measure']>[1]) =>
      [...text].reduce(
        (sum, ch) => sum + (ch.charCodeAt(0) < 128 ? style.fontSizePt * 0.52 : style.fontSizePt),
        0
      ) +
      text.length * style.characterSpacingPt;
    const mixedMeasurer: TextMeasurer = {
      ...measurer,
      hasResolvedFont: () => true,
      measure: measureMixed,
      textOutline: (text, style) => ({
        ink: { left: 0, right: measureMixed(text, style) },
      }),
    };

    const doc = part(docXml);
    const styles = part(STYLES, '/word/styles.xml').root;
    const settings = part(settingsWithNoKern, '/word/settings.xml').root;
    const theme = { ...NO_THEME_FONTS, eastAsianFontTableFaces: ['宋体'] };

    setWpsCompatMode(false);
    const wordLines = linesOf(
      layoutSemanticDocument(doc, 0, {
        measurer: mixedMeasurer,
        styleCascade: buildStyleCascadeTable(styles, theme, settings),
      })
    );
    expect(wordLines.length).toBe(3);

    setWpsCompatMode(true);
    const wpsLines = linesOf(
      layoutSemanticDocument(doc, 0, {
        measurer: mixedMeasurer,
        styleCascade: buildStyleCascadeTable(styles, theme, settings),
      })
    );
    expect(wpsLines.length).toBe(2);
  });
});
