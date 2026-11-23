import { h, type VNode } from 'vue';
import type { ToggleListPreset } from '@docx-editor.dev/core/contracts/types';
import type { ParagraphBorderEdgeWrite } from '@docx-editor.dev/core/store';
import {
  CHAR_SPACING_PRESETS,
  CN_NUMBERING_PRESETS,
  PARAGRAPH_BORDER_PRESETS,
  PARAGRAPH_SHADING_SWATCHES,
  UNDERLINE_STYLE_PRESETS,
  type UnderlineStylePreset,
} from './p1-presets';

const keepFocus = (e: MouseEvent) => e.preventDefault();

/** Clamp an open quick-bar popover horizontally and flip vertically if needed without covering its trigger. */
export function clampPopoverWithinViewport(
  bar: HTMLElement | null,
  host: HTMLElement | null
): void {
  if (!bar) return;
  const pop = bar.querySelector<HTMLElement>(
    '.weboffice-qbar__menu, .weboffice-qbar__table-grid, .weboffice-qbar__symbols-wrap'
  );
  if (!pop) return;
  pop.style.transform = '';
  pop.style.top = '';
  pop.style.bottom = '';
  pop.style.maxHeight = '';
  pop.style.overflowY = '';

  const rect = pop.getBoundingClientRect();
  const triggerRect = pop.parentElement?.getBoundingClientRect();
  const hostRect = host?.getBoundingClientRect();
  const minLeft = Math.max(8, hostRect ? Math.ceil(hostRect.left) + 8 : 8);
  const maxRight = Math.min(
    window.innerWidth - 8,
    hostRect ? Math.floor(hostRect.right) - 8 : window.innerWidth - 8
  );
  let dx = 0;
  if (rect.right > maxRight) dx -= rect.right - maxRight;
  if (rect.left + dx < minLeft) dx += minLeft - (rect.left + dx);
  if (Math.abs(dx) > 0.5) pop.style.transform = `translateX(${Math.round(dx)}px)`;

  const maxBottom = window.innerHeight - 8;
  if (rect.bottom > maxBottom && triggerRect) {
    const spaceBelow = Math.max(0, maxBottom - triggerRect.bottom - 4);
    const spaceAbove = Math.max(0, triggerRect.top - 12);
    if (spaceAbove > spaceBelow) {
      pop.style.top = 'auto';
      pop.style.bottom = 'calc(100% + 4px)';
      if (rect.height > spaceAbove) {
        pop.style.maxHeight = `${Math.max(120, Math.floor(spaceAbove))}px`;
        pop.style.overflowY = 'auto';
      }
    } else {
      pop.style.maxHeight = `${Math.max(120, Math.floor(spaceBelow))}px`;
      pop.style.overflowY = 'auto';
    }
  }
}

export function renderNumberingMenu(
  onPick: (preset: ToggleListPreset) => void,
  onRestart: () => void
): VNode {
  return h('div', { class: 'weboffice-qbar__menu weboffice-qbar__menu--num' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '中文公文编号（GB/T 9704）'),
    ...CN_NUMBERING_PRESETS.map((item) =>
      h(
        'button',
        {
          key: item.testId,
          type: 'button',
          class: 'weboffice-qbar__menu-item weboffice-qbar__num-item',
          'data-testid': item.testId,
          onMousedown: keepFocus,
          onClick: () => onPick(item.preset),
        },
        [
          h('span', { class: 'weboffice-qbar__num-sample' }, item.label),
          h('span', { class: 'weboffice-qbar__num-sub' }, item.sub),
        ]
      )
    ),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h(
      'button',
      {
        type: 'button',
        class: 'weboffice-qbar__menu-item',
        'data-testid': 'num-cn-restart',
        onMousedown: keepFocus,
        onClick: onRestart,
      },
      '重新开始编号'
    ),
    h(
      'button',
      {
        type: 'button',
        class: 'weboffice-qbar__menu-item',
        'data-testid': 'num-cn-clear',
        onMousedown: keepFocus,
        onClick: () => onPick('clear'),
      },
      '取消编号'
    ),
  ]);
}

export function renderBordersMenu(opts: {
  onApplyBorder: (
    edges: readonly ParagraphBorderEdgeWrite[],
    label: string,
    mode: 'merge' | 'replace'
  ) => void;
  onApplyShading: (hex: string | null, label: string) => void;
}): VNode {
  return h('div', { class: 'weboffice-qbar__menu weboffice-qbar__menu--borders' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '段落边框'),
    ...PARAGRAPH_BORDER_PRESETS.map((item) =>
      h(
        'button',
        {
          key: item.testId,
          type: 'button',
          class: 'weboffice-qbar__menu-item weboffice-qbar__border-item',
          'data-testid': item.testId,
          onMousedown: keepFocus,
          onClick: () => opts.onApplyBorder(item.edges, item.label, item.mode),
        },
        [
          h('span', { class: ['weboffice-qbar__border-preview', item.previewClass] }),
          h('span', null, item.label),
        ]
      )
    ),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h('div', { class: 'weboffice-qbar__section-title' }, '段落底纹'),
    h(
      'div',
      { class: 'weboffice-qbar__swatches' },
      PARAGRAPH_SHADING_SWATCHES.map((swatch) =>
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
            onClick: () => opts.onApplyShading(swatch.hex, swatch.label),
          },
          swatch.hex === null ? '无' : ''
        )
      )
    ),
  ]);
}

export function renderTextEffectsMenu(opts: {
  customSpacingPt: string;
  onUpdateCustomPt: (value: string) => void;
  onApplyUnderline: (value: UnderlineStylePreset['value'], label: string) => void;
  onApplyCharSpacing: (twips: number, label: string) => void;
  onApplyCustomSpacing: () => void;
}): VNode {
  return h('div', { class: 'weboffice-qbar__menu weboffice-qbar__menu--effects' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '下划线线型'),
    h(
      'div',
      { class: 'weboffice-qbar__u-grid' },
      UNDERLINE_STYLE_PRESETS.map((item) =>
        h(
          'button',
          {
            key: item.testId,
            type: 'button',
            class: 'weboffice-qbar__u-btn',
            'data-testid': item.testId,
            onMousedown: keepFocus,
            onClick: () => opts.onApplyUnderline(item.value, item.label),
          },
          [
            h('span', { class: ['weboffice-qbar__u-sample', item.cssClass] }, 'Aa'),
            h('span', { class: 'weboffice-qbar__u-label' }, item.label),
          ]
        )
      )
    ),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
    h('div', { class: 'weboffice-qbar__section-title' }, '字符间距'),
    h(
      'div',
      { class: 'weboffice-qbar__cspace-grid' },
      CHAR_SPACING_PRESETS.map((item) =>
        h(
          'button',
          {
            key: item.testId,
            type: 'button',
            class: 'weboffice-qbar__chip-btn',
            'data-testid': item.testId,
            onMousedown: keepFocus,
            onClick: () => opts.onApplyCharSpacing(item.twips, item.label),
          },
          item.label
        )
      )
    ),
    h('div', { class: 'weboffice-qbar__custom-row' }, [
      h('span', { class: 'weboffice-qbar__custom-label' }, '自定义'),
      h('input', {
        type: 'number',
        step: 'any',
        min: -20,
        max: 100,
        class: 'weboffice-qbar__custom-input',
        'data-testid': 'char-spacing-custom-input',
        value: opts.customSpacingPt,
        placeholder: '磅值',
        onInput: (e: Event) => opts.onUpdateCustomPt((e.target as HTMLInputElement).value),
        onKeydown: (e: KeyboardEvent) => {
          if (e.key === 'Enter' && !e.isComposing) {
            e.preventDefault();
            opts.onApplyCustomSpacing();
          }
        },
      }),
      h('span', { class: 'weboffice-qbar__custom-unit' }, '磅'),
      h(
        'button',
        {
          type: 'button',
          class: 'weboffice-qbar__custom-apply',
          'data-testid': 'char-spacing-custom-apply',
          onMousedown: keepFocus,
          onClick: opts.onApplyCustomSpacing,
        },
        '应用'
      ),
    ]),
  ]);
}

import type { TableContextQuickAction } from './qbar-p2-actions';

export function renderTableContextSection(opts: {
  canMerge: boolean;
  canToggleHeaderRow: boolean;
  isHeaderRow: boolean;
  headerRowCount: number;
  onAction: (action: TableContextQuickAction) => void;
  advancedSlot?: VNode | null;
}): VNode {
  const btn = (
    label: string,
    testId: string,
    action: TableContextQuickAction,
    extra: {
      disabled?: boolean;
      active?: boolean;
      pressed?: boolean;
      danger?: boolean;
      title?: string;
      headerRows?: number;
    } = {}
  ) =>
    h(
      'button',
      {
        type: 'button',
        class: [
          'weboffice-qbar__chip-btn',
          extra.active ? 'is-active' : '',
          extra.danger ? 'is-danger' : '',
        ],
        'data-testid': testId,
        'aria-pressed': extra.pressed !== undefined ? String(extra.pressed) : undefined,
        'data-header-rows': extra.headerRows !== undefined ? String(extra.headerRows) : undefined,
        title: extra.title,
        disabled: extra.disabled,
        onMousedown: keepFocus,
        onClick: () => opts.onAction(action),
      },
      label
    );

  const headerTitle = !opts.canToggleHeaderRow
    ? opts.headerRowCount > 0
      ? `当前前 ${opts.headerRowCount} 行为重复表头行；仅可在首行或紧邻表头的连续行切换`
      : '仅可在表格首行开启跨页重复表头'
    : opts.isHeaderRow
      ? opts.headerRowCount > 1
        ? `当前前 ${opts.headerRowCount} 行为重复表头行，点击取消当前行及后续表头`
        : '当前为重复表头行，点击取消'
      : opts.headerRowCount > 0
        ? `点击将当前行纳入重复表头（前 ${opts.headerRowCount + 1} 行）`
        : '开启首行跨页重复表头';

  return h('div', { class: 'weboffice-qbar__table-ctx', 'data-testid': 'table-ctx-actions' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '当前表格操作'),
    h('div', { class: 'weboffice-qbar__table-ctx-grid' }, [
      btn('合并单元格', 'table-ctx-merge', 'mergeCells', { disabled: !opts.canMerge }),
      btn('重复表头行', 'table-ctx-header-row', 'toggleHeaderRow', {
        disabled: !opts.canToggleHeaderRow,
        active: opts.isHeaderRow,
        pressed: opts.isHeaderRow,
        headerRows: opts.headerRowCount,
        title: headerTitle,
      }),
      btn('选中整表', 'table-ctx-select-all', 'selectTable'),
      btn('上方插行', 'table-ctx-row-above', 'rowAbove'),
      btn('下方插行', 'table-ctx-row-below', 'rowBelow'),
      btn('左侧插列', 'table-ctx-col-left', 'colLeft'),
      btn('右侧插列', 'table-ctx-col-right', 'colRight'),
      btn('删除当前行', 'table-ctx-del-row', 'deleteRow', { danger: true }),
      btn('删除当前列', 'table-ctx-del-col', 'deleteColumn', { danger: true }),
      btn('删除表格', 'table-ctx-del-table', 'deleteTable', { danger: true }),
    ]),
    opts.advancedSlot ?? null,
    h('div', { class: 'weboffice-qbar__menu-sep' }),
  ]);
}
