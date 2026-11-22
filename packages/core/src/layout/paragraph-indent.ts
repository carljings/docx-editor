// Paragraph indents from `w:ind`, bounded at the trust boundary. Split from
// paragraph-flow.ts, which re-exports these names so existing imports stay stable.

import {
  isElement,
  readTwipsMeasure,
  twips,
  twipsToPoints,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlProperty,
  type Twips,
} from '@docx-editor.dev/core/store';
import { resolveRunStyle } from './run-style.ts';
import { propertiesOfRunContainer } from './field-run-text.ts';
import { paragraphIsRtl } from './rtl-paragraph.ts';

/**
 * Soft ceiling on an indent, in twips (31_680 ≈ 22"), matching the paragraph-spacing and
 * tab-position bounds. `w:ind` is attacker-controlled and flows straight into `rightEdge`
 * and the available line width, so an unbounded value reaches paint geometry.
 */
export const MAX_PARAGRAPH_INDENT_TWIPS = 31_680;
const MAX_PARAGRAPH_INDENT_PT = twipsToPoints(twips(MAX_PARAGRAPH_INDENT_TWIPS));
const CHARS_MEASURE_PATTERN = /^[+-]?\d{1,9}(?:\.\d+)?$/;

export function indentTwips(raw: string | undefined): Twips | null {
  // The shared reader accepts up to 9 integer digits, so an oversized authored value reaches
  // the clamp rather than being read as a measurement; a longer digit string is garbage.
  const authored = readTwipsMeasure(raw);
  if (authored === null) return null;
  if (authored > MAX_PARAGRAPH_INDENT_TWIPS) return twips(MAX_PARAGRAPH_INDENT_TWIPS);
  if (authored < -MAX_PARAGRAPH_INDENT_TWIPS) return twips(-MAX_PARAGRAPH_INDENT_TWIPS);
  return twips(authored);
}

/**
 * Parse a `w:*Chars` indent attribute (hundredths of a character unit, §17.3.1.12).
 * Rejects malformed or oversized digit strings at the same 9-digit trust boundary as twips.
 */
export function indentChars(raw: string | undefined): number | null {
  if (raw === undefined) return null;
  const trimmed = raw.trim();
  if (!CHARS_MEASURE_PATTERN.test(trimmed)) return null;
  const value = Number(trimmed);
  return Number.isFinite(value) ? value : null;
}

function clampIndentPoints(points: number): number {
  if (points > MAX_PARAGRAPH_INDENT_PT) return MAX_PARAGRAPH_INDENT_PT;
  if (points < -MAX_PARAGRAPH_INDENT_PT) return -MAX_PARAGRAPH_INDENT_PT;
  return points;
}

/**
 * Resolve one twips + hundredths-of-a-char pair from a single `w:ind` element.
 *
 * Per ISO/IEC 29500-1 §17.3.1.12 and real Word/WPS behavior:
 * - A non-zero `*Chars` value supersedes the companion twips value.
 * - A zero `*Chars` paired with a non-zero twips value (authored by Word/WPS to cancel
 *   an inherited style `*Chars` while keeping an explicit twip indent) keeps the twips value.
 * - A zero `*Chars` alone (or paired with zero twips) resolves to `0`, cancelling inheritance.
 */
function resolveIndentSidePair(
  rawTwips: string | undefined,
  rawChars: string | undefined,
  charUnitPt: number
): number | null {
  const twipsVal = indentTwips(rawTwips);
  const charsVal = indentChars(rawChars);
  if (charsVal !== null && charsVal !== 0) {
    return clampIndentPoints((charsVal / 100) * charUnitPt);
  }
  if (twipsVal !== null) {
    return twipsToPoints(twipsVal);
  }
  if (charsVal === 0) {
    return 0;
  }
  return null;
}

export function paragraphIndent(
  props: readonly OoxmlProperty[],
  charUnitPt = 10
): {
  left: number;
  right: number;
} {
  let left = 0;
  let right = 0;
  const rtl = paragraphIsRtl(props);
  for (const property of props) {
    if (property.localName !== 'ind') continue;
    const attrs = property.attributes;
    // `w:left` and `w:start` are two spellings of the LEADING indent, and `w:right` and
    // `w:end` of the trailing one: in a right-to-left paragraph `w:left` indents from the
    // right margin (§17.3.1.12). The result is physical, so the sides swap there.
    const leadingTwips = attrs?.left ?? attrs?.start;
    const leadingChars = attrs?.leftChars ?? attrs?.startChars;
    const trailingTwips = attrs?.right ?? attrs?.end;
    const trailingChars = attrs?.rightChars ?? attrs?.endChars;
    const resolvedLeft = rtl
      ? resolveIndentSidePair(trailingTwips, trailingChars, charUnitPt)
      : resolveIndentSidePair(leadingTwips, leadingChars, charUnitPt);
    const resolvedRight = rtl
      ? resolveIndentSidePair(leadingTwips, leadingChars, charUnitPt)
      : resolveIndentSidePair(trailingTwips, trailingChars, charUnitPt);
    if (resolvedLeft !== null) left = resolvedLeft;
    if (resolvedRight !== null) right = resolvedRight;
  }
  return { left, right };
}

/**
 * Resolve cascaded `w:hanging` / `w:hangingChars` and `w:firstLine` / `w:firstLineChars`
 * offsets from a paragraph's cascaded property list.
 */
export function resolveParagraphFirstLineIndent(
  props: readonly OoxmlProperty[],
  charUnitPt = 10
): { hanging: number; firstLine: number } {
  let hanging = 0;
  let firstLine = 0;
  for (const property of props) {
    if (property.localName !== 'ind') continue;
    const attrs = property.attributes;
    const hasHanging = attrs?.hanging !== undefined || attrs?.hangingChars !== undefined;
    const hasFirstLine = attrs?.firstLine !== undefined || attrs?.firstLineChars !== undefined;
    // `w:hanging` and `w:firstLine` (and their `*Chars` twins) are mutually exclusive spellings
    // of one offset. Any stated attribute replaces both; otherwise an inherited hanging could
    // survive `firstLine="0"` and shift body text into the margin.
    if (!hasHanging && !hasFirstLine) continue;
    const resolvedHanging = resolveIndentSidePair(attrs?.hanging, attrs?.hangingChars, charUnitPt);
    const resolvedFirstLine = resolveIndentSidePair(
      attrs?.firstLine,
      attrs?.firstLineChars,
      charUnitPt
    );
    // `w:hanging` / `w:hangingChars` is unsigned; negative values clamp to 0.
    hanging = resolvedHanging !== null ? Math.max(0, resolvedHanging) : 0;
    // `w:firstLine` stays signed to match Word's model and the numbering reader.
    firstLine = resolvedFirstLine ?? 0;
  }
  return { hanging, firstLine };
}

const VISIBLE_RUN_CONTENT = new Set([
  't',
  'tab',
  'sym',
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'drawing',
  'pict',
  'object',
]);

const INLINE_RUN_CONTAINERS = new Set([
  'hyperlink',
  'ins',
  'sdt',
  'sdtContent',
  'smartTag',
  'fldSimple',
]);

function runDirectProperties(run: OoxmlElement): OoxmlProperty[] {
  const rPr = run.children.find(
    (child): child is OoxmlElement => isElement(child) && child.localName === 'rPr'
  );
  return propertiesOfRunContainer(rPr);
}

function hasVisibleRunContent(run: OoxmlElement): boolean {
  for (const child of run.children) {
    if (isElement(child) && VISIBLE_RUN_CONTENT.has(child.localName)) return true;
  }
  return false;
}

function findFirstVisibleRunFontSizePt(
  nodes: readonly OoxmlNode[],
  resolveRunProps: (directRunProps: readonly OoxmlProperty[]) => readonly OoxmlProperty[]
): number | null {
  for (const node of nodes) {
    if (!isElement(node)) continue;
    if (node.localName === 'r') {
      if (!hasVisibleRunContent(node)) continue;
      const direct = runDirectProperties(node);
      const style = resolveRunStyle(resolveRunProps(direct));
      if (style.hidden) continue;
      return style.fontSizePt;
    }
    if (INLINE_RUN_CONTAINERS.has(node.localName)) {
      const nested = findFirstVisibleRunFontSizePt(node.children, resolveRunProps);
      if (nested !== null) return nested;
    }
  }
  return null;
}

/**
 * Character-unit font sizes (in pt) for `w:ind/@w:*Chars` resolution:
 * - `sideCharPt` (`leftChars`/`rightChars`) follows the paragraph's cascaded mark/style font size.
 * - `firstLineCharPt` (`firstLineChars`/`hangingChars`) follows the first visible run's effective
 *   font size (matching Word/WPS so changing a paragraph's text from 16pt to 22pt widens a
 *   2-character first-line indent from 32pt to 44pt), falling back to `sideCharPt` when empty.
 */
export function resolveParagraphCharIndentUnitPt(
  paragraph: OoxmlElement,
  markRunProperties: readonly OoxmlProperty[],
  resolveRunProps: (directRunProps: readonly OoxmlProperty[]) => readonly OoxmlProperty[]
): { firstLineCharPt: number; sideCharPt: number } {
  const sideCharPt = resolveRunStyle(markRunProperties).fontSizePt;
  const firstRunPt = findFirstVisibleRunFontSizePt(paragraph.children, resolveRunProps);
  return {
    firstLineCharPt: firstRunPt ?? sideCharPt,
    sideCharPt,
  };
}

/**
 * How far a paragraph's border box reaches past its text column on each physical side.
 *
 * The box opens where the first line starts when a hanging indent moves that line outward:
 * `w:left="1701" w:hanging="1701"` boxes from the margin, not from 1701 twips. A first-line
 * indent moves the line inward and leaves the box at the indent. The first line starts on
 * the paragraph's leading side, so a right-to-left paragraph reaches past its right edge,
 * also past the margin. The other side stays at its indent.
 */
export function paragraphBoxOutdent(
  indent: { readonly hanging: number; readonly firstLine: number },
  rtl: boolean
): { readonly left: number; readonly right: number } {
  const outward = Math.max(0, indent.hanging - indent.firstLine);
  return rtl ? { left: 0, right: outward } : { left: outward, right: 0 };
}
