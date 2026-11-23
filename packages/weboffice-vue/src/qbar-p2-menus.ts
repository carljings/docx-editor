import { h, type VNode } from 'vue';
import type { Watermark, WatermarkLayout } from '@docx-editor.dev/core/contracts/types';
import { CN_OFFICIAL_SYMBOLS } from './units';
import {
  OFFICIAL_WATERMARK_PRESETS,
  TABLE_CELL_FILL_SWATCHES,
  WATERMARK_COLOR_OPTIONS,
} from './p2-presets';

const keepFocus = (e: MouseEvent) => e.preventDefault();

const LINE_SPACING_PRESETS: ReadonlyArray<{
  readonly label: string;
  readonly rule: 'exact' | 'multiple';
  readonly value: number;
  readonly testId: string;
}> = [
  { label: '固定值 28 磅（公文标准）', rule: 'exact', value: 28, testId: 'ls-exact-28' },
  { label: '固定值 29 磅', rule: 'exact', value: 29, testId: 'ls-exact-29' },
  { label: '固定值 30 磅', rule: 'exact', value: 30, testId: 'ls-exact-30' },
  { label: '固定值 22 磅', rule: 'exact', value: 22, testId: 'ls-exact-22' },
  { label: '单倍行距 (1.0×)', rule: 'multiple', value: 1, testId: 'ls-mult-1' },
  { label: '1.5 倍行距', rule: 'multiple', value: 1.5, testId: 'ls-mult-1.5' },
  { label: '2 倍行距', rule: 'multiple', value: 2, testId: 'ls-mult-2' },
];

export function renderPasteMenu(onPaste: (plainOnly: boolean) => void): VNode {
  return h('div', { class: 'weboffice-qbar__menu' }, [
    h(
      'button',
      {
        type: 'button',
        class: 'weboffice-qbar__menu-item',
        onMousedown: keepFocus,
        onClick: () => onPaste(false),
      },
      '粘贴 (Ctrl+V)'
    ),
    h(
      'button',
      {
        type: 'button',
        class: 'weboffice-qbar__menu-item',
        onMousedown: keepFocus,
        onClick: () => onPaste(true),
      },
      '仅保留文本 (Ctrl+Shift+V)'
    ),
  ]);
}

export function renderLineSpacingMenu(opts: {
  readonly onApply: (rule: 'exact' | 'multiple', value: number) => void;
  readonly onOpenDialog: () => void;
}): VNode {
  return h('div', { class: 'weboffice-qbar__menu' }, [
    ...LINE_SPACING_PRESETS.map((preset) =>
      h(
        'button',
        {
          key: preset.testId,
          type: 'button',
          class: 'weboffice-qbar__menu-item',
          'data-testid': preset.testId,
          onMousedown: keepFocus,
          onClick: () => opts.onApply(preset.rule, preset.value),
        },
        preset.label
      )
    ),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h(
      'button',
      {
        type: 'button',
        class: 'weboffice-qbar__menu-item',
        'data-testid': 'qbar-open-para-dialog',
        onMousedown: keepFocus,
        onClick: opts.onOpenDialog,
      },
      '段落排版设置（厘米 / 字符）…'
    ),
  ]);
}

export function renderTableInsertGrid(opts: {
  readonly hover: { readonly rows: number; readonly cols: number };
  readonly onHover: (rows: number, cols: number) => void;
  readonly onInsert: (rows: number, cols: number) => void;
}): VNode {
  return h('div', null, [
    h(
      'div',
      { class: 'weboffice-qbar__table-label' },
      `插入 ${opts.hover.rows} × ${opts.hover.cols} 表格`
    ),
    h(
      'div',
      { class: 'weboffice-qbar__grid-cells' },
      Array.from({ length: 6 * 8 }, (_, idx) => {
        const r = Math.floor(idx / 8) + 1;
        const c = (idx % 8) + 1;
        const active = r <= opts.hover.rows && c <= opts.hover.cols;
        return h('button', {
          key: `${r}x${c}`,
          type: 'button',
          class: ['weboffice-qbar__grid-cell', active ? 'is-on' : ''],
          'data-testid': `table-cell-${r}x${c}`,
          'aria-label': `${r} 行 ${c} 列`,
          onMousedown: keepFocus,
          onMouseenter: () => opts.onHover(r, c),
          onClick: () => opts.onInsert(r, c),
        });
      })
    ),
  ]);
}

export function renderSymbolsMenu(opts: {
  readonly footnoteEnabled: boolean;
  readonly endnoteEnabled: boolean;
  readonly textBoxEnabled: boolean;
  readonly onInsertBreak: (kind: 'page' | 'section') => void;
  readonly onInsertExtra: (kind: 'footnote' | 'endnote' | 'textBox') => void;
  readonly onInsertSymbol: (char: string) => void;
}): VNode {
  return h('div', { class: 'weboffice-qbar__symbols-wrap' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '分隔符与注释'),
    h('div', { class: 'weboffice-qbar__breaks-row' }, [
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'qbar-insert-page-break',
          onMousedown: keepFocus,
          onClick: () => opts.onInsertBreak('page'),
        },
        '插入分页符'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'qbar-insert-section-break',
          onMousedown: keepFocus,
          onClick: () => opts.onInsertBreak('section'),
        },
        '插入分节符'
      ),
    ]),
    h('div', { class: 'weboffice-qbar__valign-row' }, [
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'qbar-insert-footnote',
          disabled: !opts.footnoteEnabled,
          onMousedown: keepFocus,
          onClick: () => opts.onInsertExtra('footnote'),
        },
        '插入脚注'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'qbar-insert-endnote',
          disabled: !opts.endnoteEnabled,
          onMousedown: keepFocus,
          onClick: () => opts.onInsertExtra('endnote'),
        },
        '插入尾注'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'qbar-insert-textbox',
          disabled: !opts.textBoxEnabled,
          onMousedown: keepFocus,
          onClick: () => opts.onInsertExtra('textBox'),
        },
        '文本框'
      ),
    ]),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h('div', { class: 'weboffice-qbar__section-title' }, '公文常用符号'),
    h(
      'div',
      { class: 'weboffice-qbar__symbols-grid' },
      CN_OFFICIAL_SYMBOLS.map((sym) =>
        h(
          'button',
          {
            key: sym.char,
            type: 'button',
            class: 'weboffice-qbar__symbol-btn',
            title: sym.label,
            onMousedown: keepFocus,
            onClick: () => opts.onInsertSymbol(sym.char),
          },
          sym.char
        )
      )
    ),
  ]);
}

export type HeaderFooterQuickAction =
  | 'editHeader'
  | 'editFooter'
  | 'exitHeaderFooter'
  | 'insertOfficialPageNum'
  | 'insertPage'
  | 'insertNumPages'
  | 'insertPageXofY'
  | 'toggleFirstPage'
  | 'removeHeader'
  | 'removeFooter';

export function renderPageExtrasMenu(opts: {
  readonly editingRegion: 'header' | 'footer' | null;
  readonly titlePage: boolean;
  readonly hasWatermark: boolean;
  readonly currentWatermarkText: string;
  readonly customWmText: string;
  readonly customWmLayout: WatermarkLayout;
  readonly customWmColor: string;
  readonly onUpdateCustomText: (value: string) => void;
  readonly onUpdateCustomLayout: (value: WatermarkLayout) => void;
  readonly onUpdateCustomColor: (value: string) => void;
  readonly onHeaderFooterAction: (action: HeaderFooterQuickAction) => void;
  readonly onApplyWatermarkPreset: (watermark: Watermark, label: string) => void;
  readonly onApplyCustomWatermark: () => void;
  readonly onClearWatermark: () => void;
}): VNode {
  const hfBtn = (
    label: string,
    testId: string,
    action: HeaderFooterQuickAction,
    extra: {
      active?: boolean;
      disabled?: boolean;
      pressed?: boolean;
      accent?: boolean;
      danger?: boolean;
    } = {}
  ) =>
    h(
      'button',
      {
        type: 'button',
        class: [
          'weboffice-qbar__chip-btn',
          extra.active ? 'is-active' : '',
          extra.accent ? 'is-accent' : '',
          extra.danger ? 'is-danger' : '',
        ],
        'data-testid': testId,
        'aria-pressed': extra.pressed !== undefined ? String(extra.pressed) : undefined,
        disabled: extra.disabled,
        onMousedown: keepFocus,
        onClick: () => opts.onHeaderFooterAction(action),
      },
      label
    );

  return h('div', { class: 'weboffice-qbar__menu weboffice-qbar__menu--page-extras' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '页眉与页脚'),
    h('div', { class: 'weboffice-qbar__table-ctx-grid' }, [
      hfBtn('编辑页眉', 'hf-edit-header', 'editHeader', {
        active: opts.editingRegion === 'header',
      }),
      hfBtn('编辑页脚', 'hf-edit-footer', 'editFooter', {
        active: opts.editingRegion === 'footer',
      }),
      hfBtn('退出页眉页脚', 'hf-exit', 'exitHeaderFooter', {
        disabled: opts.editingRegion === null,
      }),
      hfBtn('首页不同', 'hf-toggle-first-page', 'toggleFirstPage', {
        active: opts.titlePage,
        pressed: opts.titlePage,
      }),
    ]),
    h(
      'div',
      { class: 'weboffice-qbar__section-title weboffice-qbar__section-title--sub' },
      '公文页码'
    ),
    h('div', { class: 'weboffice-qbar__table-ctx-grid' }, [
      hfBtn('公文页码（— 1 —）', 'hf-insert-official-page-num', 'insertOfficialPageNum', {
        accent: true,
      }),
      hfBtn('第 X 页 / 共 Y 页', 'hf-insert-x-of-y', 'insertPageXofY'),
      hfBtn('当前页码 (PAGE)', 'hf-insert-page', 'insertPage'),
      hfBtn('总页数 (NUMPAGES)', 'hf-insert-numpages', 'insertNumPages'),
      hfBtn('删除页眉', 'hf-remove-header', 'removeHeader', { danger: true }),
      hfBtn('删除页脚', 'hf-remove-footer', 'removeFooter', { danger: true }),
    ]),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h('div', { class: 'weboffice-qbar__section-row' }, [
      h('span', { class: 'weboffice-qbar__section-title' }, '公文水印'),
      opts.hasWatermark
        ? h(
            'span',
            { class: 'weboffice-qbar__wm-badge', 'data-testid': 'wm-active-badge' },
            opts.currentWatermarkText ? `当前：${opts.currentWatermarkText}` : '已设图片水印'
          )
        : null,
    ]),
    h(
      'div',
      { class: 'weboffice-qbar__table-ctx-grid' },
      OFFICIAL_WATERMARK_PRESETS.map((item) =>
        h(
          'button',
          {
            key: item.testId,
            type: 'button',
            class: 'weboffice-qbar__chip-btn',
            'data-testid': item.testId,
            onMousedown: keepFocus,
            onClick: () => opts.onApplyWatermarkPreset(item.watermark, item.label),
          },
          item.label
        )
      )
    ),
    h('div', { class: 'weboffice-qbar__wm-custom' }, [
      h('input', {
        type: 'text',
        class: 'weboffice-qbar__custom-input weboffice-qbar__wm-input',
        'data-testid': 'wm-custom-input',
        placeholder: '自定义水印文字',
        value: opts.customWmText,
        onInput: (e: Event) => opts.onUpdateCustomText((e.target as HTMLInputElement).value),
        onKeydown: (e: KeyboardEvent) => {
          if (e.key === 'Enter' && !e.isComposing) {
            e.preventDefault();
            opts.onApplyCustomWatermark();
          }
        },
      }),
      h(
        'select',
        {
          class: 'weboffice-qbar__mini-select',
          'data-testid': 'wm-custom-layout',
          'aria-label': '水印方向',
          value: opts.customWmLayout,
          onChange: (e: Event) =>
            opts.onUpdateCustomLayout((e.target as HTMLSelectElement).value as WatermarkLayout),
        },
        [h('option', { value: 'diagonal' }, '斜向'), h('option', { value: 'horizontal' }, '水平')]
      ),
      h(
        'select',
        {
          class: 'weboffice-qbar__mini-select',
          'data-testid': 'wm-custom-color',
          'aria-label': '水印颜色',
          value: opts.customWmColor,
          onChange: (e: Event) => opts.onUpdateCustomColor((e.target as HTMLSelectElement).value),
        },
        WATERMARK_COLOR_OPTIONS.map((c) => h('option', { key: c.hex, value: c.hex }, c.label))
      ),
    ]),
    h('div', { class: 'weboffice-qbar__wm-actions' }, [
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__custom-apply',
          'data-testid': 'wm-custom-apply',
          onMousedown: keepFocus,
          onClick: opts.onApplyCustomWatermark,
        },
        '应用水印'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn is-danger',
          'data-testid': 'wm-clear',
          disabled: !opts.hasWatermark,
          onMousedown: keepFocus,
          onClick: opts.onClearWatermark,
        },
        '清除水印'
      ),
    ]),
  ]);
}

export function renderTableAdvancedSection(opts: {
  readonly splitLimits: { readonly maxRows: number; readonly maxCols: number } | null;
  readonly splitRows: number;
  readonly splitCols: number;
  readonly onUpdateSplitRows: (n: number) => void;
  readonly onUpdateSplitCols: (n: number) => void;
  readonly onSplitCell: () => void;
  readonly onVerticalAlign: (alignment: 'top' | 'center' | 'bottom') => void;
  readonly onCellFill: (hex: string | null, label: string) => void;
  readonly onWholeTableBorder: (scope: 'threeLine' | 'tableAll' | 'tableNone') => void;
}): VNode {
  const canSplit = opts.splitLimits !== null;
  const maxR = opts.splitLimits?.maxRows ?? 1;
  const maxC = opts.splitLimits?.maxCols ?? 2;

  return h('div', { class: 'weboffice-qbar__table-adv' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '拆分与单元格排版'),
    h('div', { class: 'weboffice-qbar__split-row' }, [
      h('span', { class: 'weboffice-qbar__custom-label' }, '拆为'),
      h('input', {
        type: 'number',
        min: 1,
        max: maxR,
        class: 'weboffice-qbar__split-input',
        'data-testid': 'table-split-rows',
        'aria-label': '拆分行数',
        disabled: !canSplit,
        value: String(Math.max(1, Math.min(maxR, opts.splitRows))),
        onInput: (e: Event) => {
          const v = Math.round(Number((e.target as HTMLInputElement).value));
          if (Number.isFinite(v) && v >= 1) opts.onUpdateSplitRows(v);
        },
      }),
      h('span', { class: 'weboffice-qbar__custom-unit' }, '行×'),
      h('input', {
        type: 'number',
        min: 1,
        max: maxC,
        class: 'weboffice-qbar__split-input',
        'data-testid': 'table-split-cols',
        'aria-label': '拆分列数',
        disabled: !canSplit,
        value: String(Math.max(1, Math.min(maxC, opts.splitCols))),
        onInput: (e: Event) => {
          const v = Math.round(Number((e.target as HTMLInputElement).value));
          if (Number.isFinite(v) && v >= 1) opts.onUpdateSplitCols(v);
        },
      }),
      h('span', { class: 'weboffice-qbar__custom-unit' }, '列'),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-split',
          disabled: !canSplit,
          onMousedown: keepFocus,
          onClick: opts.onSplitCell,
        },
        '拆分'
      ),
    ]),
    h('div', { class: 'weboffice-qbar__valign-row' }, [
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-valign-top',
          onMousedown: keepFocus,
          onClick: () => opts.onVerticalAlign('top'),
        },
        '顶对齐'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-valign-center',
          onMousedown: keepFocus,
          onClick: () => opts.onVerticalAlign('center'),
        },
        '垂直居中'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-valign-bottom',
          onMousedown: keepFocus,
          onClick: () => opts.onVerticalAlign('bottom'),
        },
        '底对齐'
      ),
    ]),
    h(
      'div',
      { class: 'weboffice-qbar__section-title weboffice-qbar__section-title--sub' },
      '单元格底纹'
    ),
    h(
      'div',
      { class: 'weboffice-qbar__swatches' },
      TABLE_CELL_FILL_SWATCHES.map((swatch) =>
        h(
          'button',
          {
            key: swatch.testId,
            type: 'button',
            class: ['weboffice-qbar__swatch', swatch.hex === null ? 'is-clear' : ''],
            'data-testid': swatch.testId,
            title: swatch.label,
            style: swatch.hex ? { backgroundColor: `#${swatch.hex}` } : undefined,
            onMousedown: keepFocus,
            onClick: () => opts.onCellFill(swatch.hex, swatch.label),
          },
          swatch.hex === null ? '无' : ''
        )
      )
    ),
    h(
      'div',
      { class: 'weboffice-qbar__section-title weboffice-qbar__section-title--sub' },
      '整表框线'
    ),
    h('div', { class: 'weboffice-qbar__valign-row' }, [
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn is-accent',
          'data-testid': 'table-ctx-border-three-line',
          title: '应用 GB/T 7713 公文三线表（顶底 1.5 磅、表头下 0.75 磅、无竖线）',
          onMousedown: keepFocus,
          onClick: () => opts.onWholeTableBorder('threeLine'),
        },
        '公文三线表'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-border-all',
          onMousedown: keepFocus,
          onClick: () => opts.onWholeTableBorder('tableAll'),
        },
        '全框线'
      ),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__chip-btn',
          'data-testid': 'table-ctx-border-none',
          onMousedown: keepFocus,
          onClick: () => opts.onWholeTableBorder('tableNone'),
        },
        '无框线'
      ),
    ]),
  ]);
}
