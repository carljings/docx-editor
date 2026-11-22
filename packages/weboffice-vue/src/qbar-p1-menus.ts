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

export function renderNumberingMenu(onPick: (preset: ToggleListPreset) => void): VNode {
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
        'data-testid': 'num-cn-clear',
        onMousedown: keepFocus,
        onClick: () => onPick('clear'),
      },
      '取消编号'
    ),
  ]);
}

export function renderBordersMenu(opts: {
  onApplyBorder: (edges: readonly ParagraphBorderEdgeWrite[], label: string) => void;
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
          onClick: () => opts.onApplyBorder(item.edges, item.label),
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

export function renderTableContextSection(opts: {
  canMerge: boolean;
  isHeaderRow: boolean;
  onAction: (
    action:
      | 'mergeCells'
      | 'toggleHeaderRow'
      | 'rowAbove'
      | 'rowBelow'
      | 'colLeft'
      | 'colRight'
      | 'deleteRow'
      | 'deleteTable'
  ) => void;
}): VNode {
  const btn = (
    label: string,
    testId: string,
    action: Parameters<typeof opts.onAction>[0],
    extra: { disabled?: boolean; active?: boolean; danger?: boolean } = {}
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
        disabled: extra.disabled,
        onMousedown: keepFocus,
        onClick: () => opts.onAction(action),
      },
      label
    );

  return h('div', { class: 'weboffice-qbar__table-ctx', 'data-testid': 'table-ctx-actions' }, [
    h('div', { class: 'weboffice-qbar__section-title' }, '当前表格操作'),
    h('div', { class: 'weboffice-qbar__table-ctx-grid' }, [
      btn('合并单元格', 'table-ctx-merge', 'mergeCells', { disabled: !opts.canMerge }),
      btn('重复表头行', 'table-ctx-header-row', 'toggleHeaderRow', { active: opts.isHeaderRow }),
      btn('上方插行', 'table-ctx-row-above', 'rowAbove'),
      btn('下方插行', 'table-ctx-row-below', 'rowBelow'),
      btn('左侧插列', 'table-ctx-col-left', 'colLeft'),
      btn('右侧插列', 'table-ctx-col-right', 'colRight'),
      btn('删除当前行', 'table-ctx-del-row', 'deleteRow', { danger: true }),
      btn('删除表格', 'table-ctx-del-table', 'deleteTable', { danger: true }),
    ]),
    h('div', { class: 'weboffice-qbar__menu-sep' }),
  ]);
}
