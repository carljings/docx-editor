import type { ToggleListPreset } from '@docx-editor.dev/core/contracts/types';
import type { ParagraphBorderEdgeWrite } from '@docx-editor.dev/core/store';

export interface CnNumberingPresetItem {
  readonly preset: ToggleListPreset;
  readonly label: string;
  readonly sub: string;
  readonly testId: string;
}

export const CN_NUMBERING_PRESETS: readonly CnNumberingPresetItem[] = [
  {
    preset: 'cn-heading1',
    label: '一、 二、 三、',
    sub: '一级公文标题',
    testId: 'num-cn-heading1',
  },
  {
    preset: 'cn-heading2',
    label: '（一） （二） （三）',
    sub: '二级公文标题',
    testId: 'num-cn-heading2',
  },
  { preset: 'cn-decimal', label: '1. 2. 3.', sub: '三级阿拉伯数字', testId: 'num-cn-decimal' },
  { preset: 'cn-paren', label: '（1） （2） （3）', sub: '四级括号数字', testId: 'num-cn-paren' },
];

export interface ParagraphBorderPresetItem {
  readonly id: string;
  readonly label: string;
  readonly testId: string;
  readonly previewClass: string;
  readonly mode: 'merge' | 'replace';
  readonly edges: readonly ParagraphBorderEdgeWrite[];
}

const THIN_EDGE = { val: 'single', sz: '6', space: '4', color: '334155' } as const;
const BOX_EDGE = { val: 'single', sz: '6', space: '4', color: '64748B' } as const;

export const PARAGRAPH_BORDER_PRESETS: readonly ParagraphBorderPresetItem[] = [
  {
    id: 'bottom-thin',
    label: '下框线',
    testId: 'border-bottom-thin',
    previewClass: 'is-bottom-thin',
    mode: 'merge',
    edges: [{ edge: 'bottom', attributes: THIN_EDGE }],
  },
  {
    id: 'official-red',
    label: '公文红线（2.25磅）',
    testId: 'border-official-red',
    previewClass: 'is-official-red',
    mode: 'merge',
    edges: [
      { edge: 'bottom', attributes: { val: 'single', sz: '18', space: '4', color: 'DC2626' } },
    ],
  },
  {
    id: 'official-double-red',
    label: '公文双红线',
    testId: 'border-official-double-red',
    previewClass: 'is-official-double-red',
    mode: 'merge',
    edges: [
      { edge: 'bottom', attributes: { val: 'double', sz: '12', space: '4', color: 'DC2626' } },
    ],
  },
  {
    id: 'box',
    label: '外框线',
    testId: 'border-box',
    previewClass: 'is-box',
    mode: 'replace',
    edges: [
      { edge: 'top', attributes: BOX_EDGE },
      { edge: 'left', attributes: BOX_EDGE },
      { edge: 'bottom', attributes: BOX_EDGE },
      { edge: 'right', attributes: BOX_EDGE },
    ],
  },
  {
    id: 'none',
    label: '无框线',
    testId: 'border-none',
    previewClass: 'is-none',
    mode: 'replace',
    edges: [],
  },
];

export interface ParagraphShadingSwatch {
  readonly hex: string | null;
  readonly label: string;
  readonly testId: string;
}

export const PARAGRAPH_SHADING_SWATCHES: readonly ParagraphShadingSwatch[] = [
  { hex: null, label: '无底纹', testId: 'shading-none' },
  { hex: 'F1F5F9', label: '云灰', testId: 'shading-F1F5F9' },
  { hex: 'EFF6FF', label: '浅蓝', testId: 'shading-EFF6FF' },
  { hex: 'FEFCE8', label: '淡黄', testId: 'shading-FEFCE8' },
  { hex: 'F0FDF4', label: '淡绿', testId: 'shading-F0FDF4' },
  { hex: 'FEF2F2', label: '淡红', testId: 'shading-FEF2F2' },
  { hex: 'FAF5FF', label: '淡紫', testId: 'shading-FAF5FF' },
];

export interface UnderlineStylePreset {
  readonly value: 'single' | 'double' | 'thick' | 'dotted' | 'dash' | 'dotDash' | 'wave' | 'none';
  readonly label: string;
  readonly testId: string;
  readonly cssClass: string;
}

export const UNDERLINE_STYLE_PRESETS: readonly UnderlineStylePreset[] = [
  { value: 'single', label: '单下划线', testId: 'u-single', cssClass: 'is-single' },
  { value: 'double', label: '双下划线', testId: 'u-double', cssClass: 'is-double' },
  { value: 'thick', label: '粗下划线', testId: 'u-thick', cssClass: 'is-thick' },
  { value: 'dotted', label: '点线', testId: 'u-dotted', cssClass: 'is-dotted' },
  { value: 'dash', label: '虚线', testId: 'u-dash', cssClass: 'is-dash' },
  { value: 'dotDash', label: '点划线', testId: 'u-dotDash', cssClass: 'is-dotDash' },
  { value: 'wave', label: '波浪线', testId: 'u-wave', cssClass: 'is-wave' },
  { value: 'none', label: '取消下划线', testId: 'u-none', cssClass: 'is-none' },
];

export interface CharSpacingPreset {
  readonly label: string;
  readonly pt: number;
  readonly twips: number;
  readonly testId: string;
}

export const CHAR_SPACING_PRESETS: readonly CharSpacingPreset[] = [
  { label: '标准 (0 磅)', pt: 0, twips: 0, testId: 'cspace-0' },
  { label: '紧缩 (-0.5 磅)', pt: -0.5, twips: -10, testId: 'cspace-neg05' },
  { label: '略宽 (+0.5 磅)', pt: 0.5, twips: 10, testId: 'cspace-pos05' },
  { label: '加宽 (+1 磅)', pt: 1, twips: 20, testId: 'cspace-pos1' },
  { label: '标题 (+2 磅)', pt: 2, twips: 40, testId: 'cspace-pos2' },
  { label: '特宽 (+3 磅)', pt: 3, twips: 60, testId: 'cspace-pos3' },
];
