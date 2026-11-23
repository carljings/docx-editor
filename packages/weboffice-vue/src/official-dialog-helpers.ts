import { h } from 'vue';
import type { PageSetupUpdate, ParagraphFormatUpdate } from '@docx-editor.dev/vue';
import { OFFICIAL_PAPER_SIZES, charsToTwips, cmToTwips } from './units';

export type AlignDraft = '' | 'left' | 'center' | 'right' | 'justify';
export type SpecialKindDraft = '' | 'none' | 'firstLine' | 'hanging';
export type LineRuleDraft = '' | 'exact' | 'atLeast' | 'multiple';

export function parseNonNegative(raw: string, fallback = 0): number | null {
  const trimmed = raw.trim();
  if (trimmed === '') return null;
  const n = Number(trimmed);
  if (!Number.isFinite(n) || n < 0) return fallback;
  return n;
}

export interface ParagraphDraftInput {
  readonly dirty: {
    readonly align: boolean;
    readonly indentLeft: boolean;
    readonly indentRight: boolean;
    readonly special: boolean;
    readonly spaceBefore: boolean;
    readonly spaceAfter: boolean;
    readonly lineSpacing: boolean;
  };
  readonly align: AlignDraft;
  readonly indentLeftStr: string;
  readonly indentRightStr: string;
  readonly specialKind: SpecialKindDraft;
  readonly specialUnit: 'chars' | 'cm';
  readonly specialValStr: string;
  readonly spaceBeforeStr: string;
  readonly spaceAfterStr: string;
  readonly lineRule: LineRuleDraft;
  readonly lineValStr: string;
  readonly fontPt: number;
}

export function buildParagraphFormatUpdate(input: ParagraphDraftInput): {
  update: ParagraphFormatUpdate;
  hasChanges: boolean;
} {
  const update: { -readonly [K in keyof ParagraphFormatUpdate]: ParagraphFormatUpdate[K] } = {};
  let hasChanges = false;
  if (input.dirty.align && input.align !== '') {
    update.alignment = input.align;
    hasChanges = true;
  }
  if (input.dirty.indentLeft) {
    const val = parseNonNegative(input.indentLeftStr);
    if (val !== null) {
      update.indentLeftTwips = cmToTwips(Math.min(50, val));
      hasChanges = true;
    }
  }
  if (input.dirty.indentRight) {
    const val = parseNonNegative(input.indentRightStr);
    if (val !== null) {
      update.indentRightTwips = cmToTwips(Math.min(50, val));
      hasChanges = true;
    }
  }
  if (input.dirty.special && input.specialKind !== '') {
    if (input.specialKind === 'none') {
      update.indentFirstLineChars = 0;
      update.indentFirstLineTwips = 0;
      hasChanges = true;
    } else {
      const rawVal = parseNonNegative(input.specialValStr) ?? 0;
      const sign = input.specialKind === 'hanging' ? -1 : 1;
      if (input.specialUnit === 'chars') {
        update.indentFirstLineChars = sign * Math.round(Math.min(50, rawVal) * 100);
        update.indentFirstLineTwips = sign * charsToTwips(Math.min(50, rawVal), input.fontPt);
      } else {
        update.indentFirstLineTwips = sign * cmToTwips(Math.min(50, rawVal));
      }
      hasChanges = true;
    }
  }
  if (input.dirty.spaceBefore) {
    const val = parseNonNegative(input.spaceBeforeStr);
    if (val !== null) {
      update.spaceBeforePt = Math.min(1584, val);
      hasChanges = true;
    }
  }
  if (input.dirty.spaceAfter) {
    const val = parseNonNegative(input.spaceAfterStr);
    if (val !== null) {
      update.spaceAfterPt = Math.min(1584, val);
      hasChanges = true;
    }
  }
  if (input.dirty.lineSpacing && input.lineRule !== '') {
    const rawVal = Number(input.lineValStr);
    if (Number.isFinite(rawVal)) {
      update.lineSpacing = {
        rule: input.lineRule,
        value:
          input.lineRule === 'multiple'
            ? Math.max(0.5, Math.min(10, rawVal))
            : Math.max(6, Math.min(200, rawVal)),
      };
      hasChanges = true;
    }
  }
  return { update, hasChanges };
}

export interface PageDraftInput {
  readonly dirty: {
    readonly paper: boolean;
    readonly orientation: boolean;
    readonly marginTop: boolean;
    readonly marginBottom: boolean;
    readonly marginLeft: boolean;
    readonly marginRight: boolean;
    readonly columnsCount: boolean;
    readonly columnsSpace: boolean;
    readonly columnsSep: boolean;
  };
  readonly paperId: string;
  readonly customPaperSize: { readonly widthTwips: number; readonly heightTwips: number } | null;
  readonly orientation: 'portrait' | 'landscape';
  readonly marginTopStr: string;
  readonly marginBottomStr: string;
  readonly marginLeftStr: string;
  readonly marginRightStr: string;
  readonly columnsCount: number;
  readonly columnsSpaceStr: string;
  readonly columnsSeparator: boolean;
  readonly scope: 'document' | 'section';
}

export function buildPageSetupUpdate(input: PageDraftInput): {
  update: PageSetupUpdate;
  hasChanges: boolean;
  error?: string;
} {
  const pageUpdate: { -readonly [K in keyof PageSetupUpdate]: PageSetupUpdate[K] } = {};
  let hasChanges = false;
  if (input.dirty.paper) {
    if (input.paperId === 'custom' && input.customPaperSize) {
      const base = input.customPaperSize;
      pageUpdate.pageWidthTwips =
        input.orientation === 'landscape' ? base.heightTwips : base.widthTwips;
      pageUpdate.pageHeightTwips =
        input.orientation === 'landscape' ? base.widthTwips : base.heightTwips;
      pageUpdate.orientation = input.orientation;
      hasChanges = true;
    } else {
      const paper = OFFICIAL_PAPER_SIZES.find((p) => p.id === input.paperId);
      if (paper) {
        pageUpdate.pageWidthTwips =
          input.orientation === 'landscape' ? paper.heightTwips : paper.widthTwips;
        pageUpdate.pageHeightTwips =
          input.orientation === 'landscape' ? paper.widthTwips : paper.heightTwips;
        pageUpdate.orientation = input.orientation;
        hasChanges = true;
      }
    }
  } else if (input.dirty.orientation) {
    pageUpdate.orientation = input.orientation;
    hasChanges = true;
  }

  const marginChecks = [
    { dirty: input.dirty.marginTop, raw: input.marginTopStr, key: 'marginTopTwips' as const },
    {
      dirty: input.dirty.marginBottom,
      raw: input.marginBottomStr,
      key: 'marginBottomTwips' as const,
    },
    { dirty: input.dirty.marginLeft, raw: input.marginLeftStr, key: 'marginLeftTwips' as const },
    { dirty: input.dirty.marginRight, raw: input.marginRightStr, key: 'marginRightTwips' as const },
  ];
  for (const item of marginChecks) {
    if (!item.dirty) continue;
    const val = Number(item.raw);
    if (!Number.isFinite(val) || val < 0 || val > 25) {
      return {
        update: pageUpdate,
        hasChanges: false,
        error: '请输入合法的页边距数值（0 ~ 25 厘米）',
      };
    }
    pageUpdate[item.key] = cmToTwips(val);
    hasChanges = true;
  }

  if (input.dirty.columnsCount || input.dirty.columnsSpace || input.dirty.columnsSep) {
    const columnsPatch: { count?: number; spaceTwips?: number; separator?: boolean } = {};
    if (input.dirty.columnsCount) {
      columnsPatch.count = input.columnsCount;
    }
    if (input.dirty.columnsSpace && input.columnsCount > 1) {
      const spaceCm = Number(input.columnsSpaceStr);
      if (!Number.isFinite(spaceCm) || spaceCm < 0 || spaceCm > 10) {
        return {
          update: pageUpdate,
          hasChanges: false,
          error: '请输入合法的栏间距（0 ~ 10 厘米）',
        };
      }
      columnsPatch.spaceTwips = cmToTwips(spaceCm);
    }
    if (input.dirty.columnsSep && input.columnsCount > 1) {
      columnsPatch.separator = input.columnsSeparator;
    } else if (input.dirty.columnsCount && input.columnsCount === 1) {
      columnsPatch.separator = false;
    }
    if (Object.keys(columnsPatch).length > 0) {
      pageUpdate.columns = columnsPatch;
      hasChanges = true;
    }
  }

  if (hasChanges) {
    pageUpdate.scope = input.scope;
  }
  return { update: pageUpdate, hasChanges };
}

export function renderDialogNumInput(
  label: string,
  value: string,
  onChange: (next: string) => void,
  opts: {
    min?: number;
    max?: number;
    unit?: string;
    testId?: string;
    placeholder?: string;
  } = {}
) {
  return h('label', { class: 'weboffice-dialog__field' }, [
    h('span', { class: 'weboffice-dialog__label' }, label),
    h('span', { class: 'weboffice-dialog__input-wrap' }, [
      h('input', {
        type: 'number',
        inputmode: 'decimal',
        class: 'weboffice-dialog__input',
        'data-testid': opts.testId,
        step: 'any',
        min: opts.min ?? 0,
        max: opts.max ?? 50,
        placeholder: opts.placeholder,
        value,
        onInput: (e: Event) => {
          onChange((e.target as HTMLInputElement).value);
        },
      }),
      opts.unit ? h('span', { class: 'weboffice-dialog__unit' }, opts.unit) : null,
    ]),
  ]);
}
