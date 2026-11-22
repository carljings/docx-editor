/**
 * 公文排版度量单位、中文字号步进、字数统计与常用公文符号（纯函数模块，无 DOM 依赖）。
 */

import { CN_FONT_SIZES, type ChineseFontSizeOption } from './fonts';

/** OOXML 标准：1 英寸 = 1440 twips，1 英寸 = 2.54 厘米 */
export const TWIPS_PER_INCH = 1440;
export const CM_PER_INCH = 2.54;
export const TWIPS_PER_CM = TWIPS_PER_INCH / CM_PER_INCH;
export const TWIPS_PER_PT = 20;

export interface OfficialMarginsCm {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

/** GB/T 9704-2012 党政机关公文格式国家标准页边距（厘米） */
export const GB_T_9704_MARGINS_CM: Readonly<OfficialMarginsCm> = Object.freeze({
  top: 3.7,
  bottom: 3.5,
  left: 2.8,
  right: 2.6,
});

export interface OfficialPaperSizePreset {
  readonly id: string;
  readonly label: string;
  readonly widthTwips: number;
  readonly heightTwips: number;
}

export const OFFICIAL_PAPER_SIZES: readonly OfficialPaperSizePreset[] = Object.freeze([
  { id: 'a4', label: 'A4 (21.0 × 29.7 cm)', widthTwips: 11906, heightTwips: 16838 },
  { id: 'a3', label: 'A3 (29.7 × 42.0 cm)', widthTwips: 16838, heightTwips: 23811 },
  { id: 'b5', label: 'B5 (17.6 × 25.0 cm)', widthTwips: 9979, heightTwips: 14175 },
  { id: '16k', label: '16开 (18.4 × 26.0 cm)', widthTwips: 10433, heightTwips: 14742 },
  { id: 'a5', label: 'A5 (14.8 × 21.0 cm)', widthTwips: 8391, heightTwips: 11906 },
  { id: 'letter', label: '信纸 Letter (21.6 × 27.9 cm)', widthTwips: 12240, heightTwips: 15840 },
]);

export function twipsToCm(twips: number): number {
  return Math.round((twips / TWIPS_PER_CM) * 100) / 100;
}

export function cmToTwips(cm: number): number {
  return Math.round(cm * TWIPS_PER_CM);
}

export function twipsToPt(twips: number): number {
  return Math.round((twips / TWIPS_PER_PT) * 100) / 100;
}

export function ptToTwips(pt: number): number {
  return Math.round(pt * TWIPS_PER_PT);
}

/**
 * 将“字符数”换算为 twips（以当前字号 1 全角字符 = fontSizePt pt 为基准）。
 * 例如三号字 (16pt) 首行缩进 2 字符 = 32pt = 640 twips。
 */
export function charsToTwips(chars: number, fontSizePt = 16): number {
  const safePt = Number.isFinite(fontSizePt) && fontSizePt > 0 ? fontSizePt : 16;
  return Math.round(chars * safePt * TWIPS_PER_PT);
}

export function twipsToChars(twips: number, fontSizePt = 16): number {
  const safePt = Number.isFinite(fontSizePt) && fontSizePt > 0 ? fontSizePt : 16;
  return Math.round((twips / (safePt * TWIPS_PER_PT)) * 100) / 100;
}

/**
 * 按 GB/T 9704-2012 中文字号阶梯（八号 5pt ～ 初号 42pt）升降一档，返回目标半磅值 (`w:sz`)。
 * 若已到达最大/最小档位则返回 `null`。
 */
export function stepChineseFontSize(
  currentHalfPoints: number | null,
  direction: 'increase' | 'decrease',
  sizes: readonly ChineseFontSizeOption[] = CN_FONT_SIZES
): number | null {
  if (!sizes.length) return null;
  const ascending = [...sizes]
    .map((item) => Math.round(item.pt * 2))
    .sort((a, b) => a - b)
    .filter((v, idx, arr) => idx === 0 || arr[idx - 1] !== v);
  if (currentHalfPoints === null || !Number.isFinite(currentHalfPoints)) {
    // 无有效当前字号时，以公文默认三号（16pt = 32 half-points）为基准步进
    return direction === 'increase' ? 36 : 28;
  }
  if (direction === 'increase') {
    const next = ascending.find((hp) => hp > currentHalfPoints);
    return next ?? null;
  }
  for (let i = ascending.length - 1; i >= 0; i -= 1) {
    if (ascending[i]! < currentHalfPoints) return ascending[i]!;
  }
  return null;
}

export interface DocumentWordStats {
  /** WPS/Word 常用“字数”口径：中文字符数 + 连续英文/数字词数 + 非空白标点（这里主显非空白字符总数） */
  readonly totalCharsNoSpace: number;
  readonly totalCharsWithSpace: number;
  readonly cjkChars: number;
  readonly latinWords: number;
  readonly paragraphCount: number;
}

const CJK_CHAR_RE = /[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef]/g;
const LATIN_WORD_RE = /[A-Za-z0-9]+(?:['’-][A-Za-z0-9]+)*/g;
const NON_SPACE_RE = /\S/g;

/**
 * 统计文档字数信息（兼容中英文公文统计习惯）。
 */
export function countDocumentWords(
  paragraphs: ReadonlyArray<{ readonly text: string }>
): DocumentWordStats {
  let totalCharsNoSpace = 0;
  let totalCharsWithSpace = 0;
  let cjkChars = 0;
  let latinWords = 0;
  for (const para of paragraphs) {
    const text = para.text ?? '';
    totalCharsWithSpace += text.length;
    const nonSpace = text.match(NON_SPACE_RE);
    if (nonSpace) totalCharsNoSpace += nonSpace.length;
    const cjk = text.match(CJK_CHAR_RE);
    if (cjk) cjkChars += cjk.length;
    const words = text.match(LATIN_WORD_RE);
    if (words) latinWords += words.length;
  }
  return {
    totalCharsNoSpace,
    totalCharsWithSpace,
    cjkChars,
    latinWords,
    paragraphCount: paragraphs.length,
  };
}

export interface OfficialSymbolOption {
  readonly char: string;
  readonly label: string;
}

/** 政务公文高频符号快捷清单 */
export const CN_OFFICIAL_SYMBOLS: readonly OfficialSymbolOption[] = Object.freeze([
  { char: '〔', label: '六角左括号（发文字号）' },
  { char: '〕', label: '六角右括号（发文字号）' },
  { char: '—', label: '一字线（页码/破折）' },
  { char: '……', label: '中文省略号' },
  { char: '·', label: '间隔号' },
  { char: '《', label: '左书名号' },
  { char: '》', label: '右书名号' },
  { char: '【', label: '左方头括号' },
  { char: '】', label: '右方头括号' },
  { char: '℃', label: '摄氏度' },
  { char: '‰', label: '千分号' },
  { char: '☑', label: '勾选框' },
  { char: '☐', label: '空白方框' },
  { char: '○', label: '空心圆' },
  { char: '●', label: '实心圆' },
  { char: '★', label: '五角星' },
  { char: '§', label: '章节号' },
  { char: '±', label: '正负号' },
  { char: '×', label: '乘号' },
  { char: '÷', label: '除号' },
  { char: '≥', label: '大于等于' },
  { char: '≤', label: '小于等于' },
  { char: '①', label: '圆圈一' },
  { char: '②', label: '圆圈二' },
  { char: '③', label: '圆圈三' },
  { char: '④', label: '圆圈四' },
  { char: '⑤', label: '圆圈五' },
]);
