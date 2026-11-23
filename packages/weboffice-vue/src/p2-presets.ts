import type { Watermark, WatermarkLayout } from '@docx-editor.dev/core/contracts/types';

export interface OfficialWatermarkPreset {
  readonly id: string;
  readonly label: string;
  readonly testId: string;
  readonly watermark: Watermark;
}

export const OFFICIAL_WATERMARK_PRESETS: readonly OfficialWatermarkPreset[] = [
  {
    id: 'internal',
    label: '内部资料',
    testId: 'wm-preset-internal',
    watermark: {
      text: '内部资料',
      fontFamily: '宋体',
      color: 'C0C0C0',
      transparency: 0.5,
      layout: 'diagonal',
    },
  },
  {
    id: 'confidential',
    label: '机密',
    testId: 'wm-preset-confidential',
    watermark: {
      text: '机密',
      fontFamily: '黑体',
      color: 'DC2626',
      transparency: 0.65,
      layout: 'diagonal',
    },
  },
  {
    id: 'draft',
    label: '草稿 · 仅供核稿',
    testId: 'wm-preset-draft',
    watermark: {
      text: '草稿 · 仅供核稿',
      fontFamily: '宋体',
      color: '94A3B8',
      transparency: 0.5,
      layout: 'diagonal',
    },
  },
  {
    id: 'noforward',
    label: '严禁外传',
    testId: 'wm-preset-noforward',
    watermark: {
      text: '严禁外传',
      fontFamily: '黑体',
      color: 'C0C0C0',
      transparency: 0.5,
      layout: 'diagonal',
    },
  },
];

export const WATERMARK_COLOR_OPTIONS: ReadonlyArray<{
  readonly hex: string;
  readonly label: string;
}> = [
  { hex: 'C0C0C0', label: '浅灰' },
  { hex: '94A3B8', label: '深灰' },
  { hex: 'DC2626', label: '公文红' },
  { hex: '2563EB', label: '公务蓝' },
];

export interface TableCellFillSwatch {
  readonly hex: string | null;
  readonly label: string;
  readonly testId: string;
}

export const TABLE_CELL_FILL_SWATCHES: readonly TableCellFillSwatch[] = [
  { hex: null, label: '无填充', testId: 'table-ctx-fill-none' },
  { hex: 'F1F5F9', label: '云灰', testId: 'table-ctx-fill-F1F5F9' },
  { hex: 'E2E8F0', label: '公文灰', testId: 'table-ctx-fill-E2E8F0' },
  { hex: 'EFF6FF', label: '浅蓝', testId: 'table-ctx-fill-EFF6FF' },
  { hex: 'FEFCE8', label: '淡黄', testId: 'table-ctx-fill-FEFCE8' },
  { hex: 'F0FDF4', label: '淡绿', testId: 'table-ctx-fill-F0FDF4' },
  { hex: 'FEF2F2', label: '淡红', testId: 'table-ctx-fill-FEF2F2' },
];

export function buildCustomWatermark(input: {
  readonly text: string;
  readonly layout: WatermarkLayout;
  readonly color: string;
  readonly transparency?: number;
}):
  | { readonly ok: true; readonly watermark: Watermark }
  | { readonly ok: false; readonly error: string } {
  const trimmed = input.text.replace(/[\r\n\v\f\u2028\u2029\s]+/gu, ' ').trim();
  if (!trimmed) return { ok: false, error: '请输入水印文字' };
  if (trimmed.length > 64) return { ok: false, error: '水印文字建议不超过 64 个字符' };
  const hex = /^[0-9a-fA-F]{6}$/.test(input.color) ? input.color.toUpperCase() : 'C0C0C0';
  const transparency =
    typeof input.transparency === 'number' && Number.isFinite(input.transparency)
      ? Math.max(0.1, Math.min(0.9, input.transparency))
      : 0.5;
  return {
    ok: true,
    watermark: {
      text: trimmed,
      fontFamily: '宋体',
      color: hex,
      transparency,
      layout: input.layout === 'horizontal' ? 'horizontal' : 'diagonal',
    },
  };
}

export interface FallbackOfficialHeading {
  readonly blockId: string;
  readonly text: string;
  readonly depth: 0 | 1;
}

const CN_LEVEL1_RE = /^[一二三四五六七八九十百]+[、．.]/u;
const CN_LEVEL2_RE = /^[（(][一二三四五六七八九十]+[）)]/u;

/**
 * Fallback outline scanner for Chinese official documents that use plain/numbered headings
 * (`一、...`, `（一）...`) without Word Heading styles or `w:outlineLvl`.
 */
export function extractOfficialFallbackHeadings(
  paragraphs: ReadonlyArray<{ readonly blockId: string; readonly text: string }>
): readonly FallbackOfficialHeading[] {
  const out: FallbackOfficialHeading[] = [];
  for (const para of paragraphs) {
    const clean = para.text.trim();
    if (!clean || clean.length > 64) continue;
    if (CN_LEVEL1_RE.test(clean)) {
      out.push({ blockId: para.blockId, text: clean, depth: 0 });
    } else if (CN_LEVEL2_RE.test(clean)) {
      out.push({ blockId: para.blockId, text: clean, depth: 1 });
    }
    if (out.length >= 120) break;
  }
  return out;
}
