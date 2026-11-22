import { Teleport, defineComponent, h, nextTick, ref, watch, type PropType } from 'vue';
import {
  useDocxEditor,
  useEditorState,
  usePageSetup,
  useParagraphFormat,
} from '@docx-editor.dev/vue';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import {
  GB_T_9704_MARGINS_CM,
  OFFICIAL_PAPER_SIZES,
  charsToTwips,
  cmToTwips,
  twipsToChars,
  twipsToCm,
} from './units';

export type OfficialFormatTab = 'paragraph' | 'page';

const selectCurrentFontSizePt = (snapshot: EditorSnapshot): number => {
  const pt = snapshot.formatting?.fontSizePt;
  return typeof pt === 'number' && pt > 0 ? pt : 16;
};

/**
 * 中文公文排版设置对话框（厘米 / 字符 / 磅）：
 * - 段落排版页签：对齐、左右缩进（厘米）、首行/悬挂缩进（支持“字符”与“厘米”切换）、段前段后（磅）、行距（固定值/最小值/多倍）；
 * - 页面设置页签：纸张（A4/A3/B5/16开/A5）、纵横向、页边距（厘米），内置 GB/T 9704-2012 公文标准页边距一键预设；
 * - 提交时调用 Core 原生 `setParagraphFormat` / `setPageSetup`，原子单步可撤销。
 */
export const OfficialFormatDialog = defineComponent({
  name: 'OfficialFormatDialog',
  props: {
    open: {
      type: Boolean,
      required: true,
    },
    initialTab: {
      type: String as PropType<OfficialFormatTab>,
      default: 'paragraph',
    },
    hostElement: {
      type: Object as PropType<HTMLElement | null>,
      default: null,
    },
    onClose: {
      type: Function as PropType<() => void>,
      required: true,
    },
  },
  setup(props) {
    // oxlint-disable-next-line react/rules-of-hooks
    const editorRef = useDocxEditor();
    // oxlint-disable-next-line react/rules-of-hooks
    const paragraph = useParagraphFormat();
    // oxlint-disable-next-line react/rules-of-hooks
    const pageSetup = usePageSetup();
    // oxlint-disable-next-line react/rules-of-hooks
    const currentFontSizePt = useEditorState(selectCurrentFontSizePt);

    const dialogRef = ref<HTMLDialogElement | null>(null);
    const activeTab = ref<OfficialFormatTab>('paragraph');
    const errorText = ref('');

    // 段落草稿态
    const align = ref<'left' | 'center' | 'right' | 'justify'>('left');
    const indentLeftCm = ref(0);
    const indentRightCm = ref(0);
    const specialKind = ref<'none' | 'firstLine' | 'hanging'>('none');
    const specialUnit = ref<'chars' | 'cm'>('chars');
    const specialValue = ref(2);
    const spaceBeforePt = ref(0);
    const spaceAfterPt = ref(0);
    const lineRule = ref<'exact' | 'atLeast' | 'multiple'>('exact');
    const lineValue = ref(28);

    // 页面草稿态
    const paperId = ref('a4');
    const orientation = ref<'portrait' | 'landscape'>('portrait');
    const marginTopCm = ref(GB_T_9704_MARGINS_CM.top);
    const marginBottomCm = ref(GB_T_9704_MARGINS_CM.bottom);
    const marginLeftCm = ref(GB_T_9704_MARGINS_CM.left);
    const marginRightCm = ref(GB_T_9704_MARGINS_CM.right);
    const pageScope = ref<'document' | 'section'>('document');

    function seedFromEditor() {
      errorText.value = '';
      // Seed paragraph fields
      const pf = paragraph.format.value;
      const fontPt = currentFontSizePt.value;
      if (pf) {
        align.value = pf.alignment ?? 'left';
        indentLeftCm.value = twipsToCm(pf.indentLeftTwips ?? 0);
        indentRightCm.value = twipsToCm(pf.indentRightTwips ?? 0);
        const fl = pf.indentFirstLineTwips ?? 0;
        if (fl > 0) {
          specialKind.value = 'firstLine';
          specialUnit.value = 'chars';
          specialValue.value = twipsToChars(fl, fontPt);
        } else if (fl < 0) {
          specialKind.value = 'hanging';
          specialUnit.value = 'cm';
          specialValue.value = twipsToCm(Math.abs(fl));
        } else {
          specialKind.value = 'none';
          specialUnit.value = 'chars';
          specialValue.value = 2;
        }
        spaceBeforePt.value = Math.round((pf.spaceBeforePt ?? 0) * 10) / 10;
        spaceAfterPt.value = Math.round((pf.spaceAfterPt ?? 0) * 10) / 10;
        if (pf.lineSpacing) {
          lineRule.value = pf.lineSpacing.rule;
          lineValue.value = Math.round(pf.lineSpacing.value * 100) / 100;
        } else {
          lineRule.value = 'exact';
          lineValue.value = 28;
        }
      }

      // Seed page setup fields
      const ps = pageSetup.pageSetup.value;
      if (ps) {
        orientation.value = ps.orientation;
        const shortSide = Math.min(ps.pageWidthTwips, ps.pageHeightTwips);
        const longSide = Math.max(ps.pageWidthTwips, ps.pageHeightTwips);
        const matched = OFFICIAL_PAPER_SIZES.find(
          (p) => Math.abs(p.widthTwips - shortSide) < 40 && Math.abs(p.heightTwips - longSide) < 40
        );
        paperId.value = matched?.id ?? 'a4';
        marginTopCm.value = twipsToCm(ps.marginsTwips.top);
        marginBottomCm.value = twipsToCm(ps.marginsTwips.bottom);
        marginLeftCm.value = twipsToCm(ps.marginsTwips.left);
        marginRightCm.value = twipsToCm(ps.marginsTwips.right);
      }
    }

    watch(
      () => props.open,
      async (isOpen) => {
        if (isOpen) {
          activeTab.value = props.initialTab;
          seedFromEditor();
          await nextTick();
          if (dialogRef.value && !dialogRef.value.open) {
            dialogRef.value.showModal();
          }
        } else if (dialogRef.value?.open) {
          dialogRef.value.close();
        }
      }
    );

    function switchSpecialUnit(nextUnit: 'chars' | 'cm') {
      if (specialUnit.value === nextUnit) return;
      const fontPt = currentFontSizePt.value;
      const currentTwips =
        specialUnit.value === 'chars'
          ? charsToTwips(specialValue.value, fontPt)
          : cmToTwips(specialValue.value);
      specialUnit.value = nextUnit;
      specialValue.value =
        nextUnit === 'chars' ? twipsToChars(currentTwips, fontPt) : twipsToCm(currentTwips);
    }

    function applyOfficialBodyPreset() {
      align.value = 'justify';
      indentLeftCm.value = 0;
      indentRightCm.value = 0;
      specialKind.value = 'firstLine';
      specialUnit.value = 'chars';
      specialValue.value = 2;
      spaceBeforePt.value = 0;
      spaceAfterPt.value = 0;
      lineRule.value = 'exact';
      lineValue.value = 28;
    }

    function applyOfficialMarginPreset() {
      paperId.value = 'a4';
      orientation.value = 'portrait';
      marginTopCm.value = GB_T_9704_MARGINS_CM.top;
      marginBottomCm.value = GB_T_9704_MARGINS_CM.bottom;
      marginLeftCm.value = GB_T_9704_MARGINS_CM.left;
      marginRightCm.value = GB_T_9704_MARGINS_CM.right;
    }

    function applyStandardMarginPreset() {
      marginTopCm.value = 2.54;
      marginBottomCm.value = 2.54;
      marginLeftCm.value = 3.18;
      marginRightCm.value = 3.18;
    }

    function handleClose() {
      if (dialogRef.value?.open) dialogRef.value.close();
      props.onClose();
      editorRef.value?.focus();
    }

    function handleSubmit(e: Event) {
      e.preventDefault();
      errorText.value = '';
      if (activeTab.value === 'paragraph') {
        const fontPt = currentFontSizePt.value;
        const magTwips =
          specialKind.value === 'none'
            ? 0
            : specialUnit.value === 'chars'
              ? charsToTwips(Math.max(0, specialValue.value), fontPt)
              : cmToTwips(Math.max(0, specialValue.value));
        const signedFirstLine =
          specialKind.value === 'hanging'
            ? -magTwips
            : specialKind.value === 'firstLine'
              ? magTwips
              : 0;
        const ok = paragraph.apply({
          alignment: align.value,
          indentLeftTwips: cmToTwips(Math.max(0, indentLeftCm.value)),
          indentRightTwips: cmToTwips(Math.max(0, indentRightCm.value)),
          indentFirstLineTwips: signedFirstLine,
          spaceBeforePt: Math.max(0, spaceBeforePt.value),
          spaceAfterPt: Math.max(0, spaceAfterPt.value),
          lineSpacing: {
            rule: lineRule.value,
            value:
              lineRule.value === 'multiple'
                ? Math.max(0.5, Math.min(10, lineValue.value))
                : Math.max(6, Math.min(200, lineValue.value)),
          },
        });
        if (!ok) {
          errorText.value = '当前位置无法修改段落格式（可能为只读或受保护区域）';
          return;
        }
        handleClose();
        return;
      }

      const paper =
        OFFICIAL_PAPER_SIZES.find((p) => p.id === paperId.value) ?? OFFICIAL_PAPER_SIZES[0]!;
      const pageWidthTwips =
        orientation.value === 'landscape' ? paper.heightTwips : paper.widthTwips;
      const pageHeightTwips =
        orientation.value === 'landscape' ? paper.widthTwips : paper.heightTwips;
      const ok = pageSetup.apply({
        pageWidthTwips,
        pageHeightTwips,
        orientation: orientation.value,
        marginTopTwips: cmToTwips(Math.max(0.5, marginTopCm.value)),
        marginBottomTwips: cmToTwips(Math.max(0.5, marginBottomCm.value)),
        marginLeftTwips: cmToTwips(Math.max(0.5, marginLeftCm.value)),
        marginRightTwips: cmToTwips(Math.max(0.5, marginRightCm.value)),
        scope: pageScope.value,
      });
      if (!ok) {
        errorText.value = '当前文档无法应用页面设置（请检查页边距是否过大或处于只读模式）';
        return;
      }
      handleClose();
    }

    const numInput = (
      label: string,
      value: number,
      onChange: (n: number) => void,
      opts: { step?: number; min?: number; max?: number; unit?: string; testId?: string } = {}
    ) =>
      h('label', { class: 'weboffice-dialog__field' }, [
        h('span', { class: 'weboffice-dialog__label' }, label),
        h('span', { class: 'weboffice-dialog__input-wrap' }, [
          h('input', {
            type: 'number',
            class: 'weboffice-dialog__input',
            'data-testid': opts.testId,
            step: opts.step ?? 0.1,
            min: opts.min ?? 0,
            max: opts.max ?? 50,
            value: String(value),
            onInput: (e: Event) => {
              const parsed = Number((e.target as HTMLInputElement).value);
              if (Number.isFinite(parsed)) onChange(parsed);
            },
          }),
          opts.unit ? h('span', { class: 'weboffice-dialog__unit' }, opts.unit) : null,
        ]),
      ]);

    return () => {
      if (!props.open || !props.hostElement) return null;
      return h(Teleport, { to: props.hostElement }, [
        h(
          'dialog',
          {
            ref: dialogRef,
            class: 'weboffice-dialog',
            'data-testid': 'weboffice-official-format-dialog',
            onCancel: (e: Event) => {
              e.preventDefault();
              handleClose();
            },
          },
          [
            h('form', { class: 'weboffice-dialog__form', onSubmit: handleSubmit }, [
              h('header', { class: 'weboffice-dialog__head' }, [
                h('div', { class: 'weboffice-dialog__tabs', role: 'tablist' }, [
                  h(
                    'button',
                    {
                      type: 'button',
                      role: 'tab',
                      'aria-selected': activeTab.value === 'paragraph',
                      class: [
                        'weboffice-dialog__tab',
                        activeTab.value === 'paragraph' ? 'is-active' : '',
                      ],
                      'data-testid': 'official-tab-paragraph',
                      onClick: () => {
                        activeTab.value = 'paragraph';
                        errorText.value = '';
                      },
                    },
                    '段落排版（厘米 / 字符）'
                  ),
                  h(
                    'button',
                    {
                      type: 'button',
                      role: 'tab',
                      'aria-selected': activeTab.value === 'page',
                      class: [
                        'weboffice-dialog__tab',
                        activeTab.value === 'page' ? 'is-active' : '',
                      ],
                      'data-testid': 'official-tab-page',
                      onClick: () => {
                        activeTab.value = 'page';
                        errorText.value = '';
                      },
                    },
                    '页面设置（厘米）'
                  ),
                ]),
                h(
                  'button',
                  {
                    type: 'button',
                    class: 'weboffice-dialog__close',
                    'aria-label': '关闭',
                    onClick: handleClose,
                  },
                  '×'
                ),
              ]),
              activeTab.value === 'paragraph'
                ? h('div', { class: 'weboffice-dialog__body' }, [
                    h('div', { class: 'weboffice-dialog__presets' }, [
                      h('span', { class: 'weboffice-dialog__preset-label' }, '快捷预设：'),
                      h(
                        'button',
                        {
                          type: 'button',
                          class: 'weboffice-dialog__chip',
                          'data-testid': 'preset-official-body',
                          onClick: applyOfficialBodyPreset,
                        },
                        '公文正文（首行2字符 · 固定28磅）'
                      ),
                    ]),
                    h('div', { class: 'weboffice-dialog__grid' }, [
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '对齐方式'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            value: align.value,
                            onChange: (e: Event) => {
                              align.value = (e.target as HTMLSelectElement)
                                .value as typeof align.value;
                            },
                          },
                          [
                            h('option', { value: 'left' }, '左对齐'),
                            h('option', { value: 'center' }, '居中对齐'),
                            h('option', { value: 'right' }, '右对齐'),
                            h('option', { value: 'justify' }, '两端对齐'),
                          ]
                        ),
                      ]),
                      h(
                        'div',
                        { class: 'weboffice-dialog__hint' },
                        `当前字号参考：${currentFontSizePt.value} 磅`
                      ),
                      numInput('左缩进', indentLeftCm.value, (n) => (indentLeftCm.value = n), {
                        unit: '厘米',
                        step: 0.1,
                      }),
                      numInput('右缩进', indentRightCm.value, (n) => (indentRightCm.value = n), {
                        unit: '厘米',
                        step: 0.1,
                      }),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '特殊缩进'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            'data-testid': 'para-special-kind',
                            value: specialKind.value,
                            onChange: (e: Event) => {
                              specialKind.value = (e.target as HTMLSelectElement)
                                .value as typeof specialKind.value;
                            },
                          },
                          [
                            h('option', { value: 'none' }, '无'),
                            h('option', { value: 'firstLine' }, '首行缩进'),
                            h('option', { value: 'hanging' }, '悬挂缩进'),
                          ]
                        ),
                      ]),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '缩进度量值'),
                        h('span', { class: 'weboffice-dialog__input-wrap' }, [
                          h('input', {
                            type: 'number',
                            class: 'weboffice-dialog__input',
                            'data-testid': 'para-special-val',
                            disabled: specialKind.value === 'none',
                            step: specialUnit.value === 'chars' ? 0.5 : 0.1,
                            min: 0,
                            max: 20,
                            value: String(specialValue.value),
                            onInput: (e: Event) => {
                              const parsed = Number((e.target as HTMLInputElement).value);
                              if (Number.isFinite(parsed)) specialValue.value = parsed;
                            },
                          }),
                          h(
                            'select',
                            {
                              class: 'weboffice-dialog__unit-select',
                              disabled: specialKind.value === 'none',
                              value: specialUnit.value,
                              onChange: (e: Event) =>
                                switchSpecialUnit(
                                  (e.target as HTMLSelectElement).value as 'chars' | 'cm'
                                ),
                            },
                            [
                              h('option', { value: 'chars' }, '字符'),
                              h('option', { value: 'cm' }, '厘米'),
                            ]
                          ),
                        ]),
                      ]),
                      numInput('段前间距', spaceBeforePt.value, (n) => (spaceBeforePt.value = n), {
                        unit: '磅',
                        step: 1,
                        max: 200,
                      }),
                      numInput('段后间距', spaceAfterPt.value, (n) => (spaceAfterPt.value = n), {
                        unit: '磅',
                        step: 1,
                        max: 200,
                      }),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '行距规则'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            'data-testid': 'para-line-rule',
                            value: lineRule.value,
                            onChange: (e: Event) => {
                              const nextRule = (e.target as HTMLSelectElement)
                                .value as typeof lineRule.value;
                              if (nextRule === 'multiple' && lineRule.value !== 'multiple') {
                                lineValue.value = 1.5;
                              } else if (nextRule !== 'multiple' && lineRule.value === 'multiple') {
                                lineValue.value = 28;
                              }
                              lineRule.value = nextRule;
                            },
                          },
                          [
                            h('option', { value: 'exact' }, '固定值（公文推荐）'),
                            h('option', { value: 'atLeast' }, '最小值'),
                            h('option', { value: 'multiple' }, '多倍行距'),
                          ]
                        ),
                      ]),
                      numInput('行距设置值', lineValue.value, (n) => (lineValue.value = n), {
                        unit: lineRule.value === 'multiple' ? '倍' : '磅',
                        step: lineRule.value === 'multiple' ? 0.05 : 1,
                        min: lineRule.value === 'multiple' ? 0.5 : 6,
                        max: lineRule.value === 'multiple' ? 10 : 200,
                        testId: 'para-line-val',
                      }),
                    ]),
                  ])
                : h('div', { class: 'weboffice-dialog__body' }, [
                    h('div', { class: 'weboffice-dialog__presets' }, [
                      h('span', { class: 'weboffice-dialog__preset-label' }, '标准边距：'),
                      h(
                        'button',
                        {
                          type: 'button',
                          class: 'weboffice-dialog__chip',
                          'data-testid': 'preset-gb9704-margins',
                          onClick: applyOfficialMarginPreset,
                        },
                        'GB/T 9704 公文（3.7/3.5/2.8/2.6 cm）'
                      ),
                      h(
                        'button',
                        {
                          type: 'button',
                          class: 'weboffice-dialog__chip',
                          onClick: applyStandardMarginPreset,
                        },
                        '常规（2.54/3.18 cm）'
                      ),
                    ]),
                    h('div', { class: 'weboffice-dialog__grid' }, [
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '纸张大小'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            value: paperId.value,
                            onChange: (e: Event) => {
                              paperId.value = (e.target as HTMLSelectElement).value;
                            },
                          },
                          OFFICIAL_PAPER_SIZES.map((p) =>
                            h('option', { key: p.id, value: p.id }, p.label)
                          )
                        ),
                      ]),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '纸张方向'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            value: orientation.value,
                            onChange: (e: Event) => {
                              orientation.value = (e.target as HTMLSelectElement).value as
                                'portrait' | 'landscape';
                            },
                          },
                          [
                            h('option', { value: 'portrait' }, '纵向'),
                            h('option', { value: 'landscape' }, '横向'),
                          ]
                        ),
                      ]),
                      numInput('上边距', marginTopCm.value, (n) => (marginTopCm.value = n), {
                        unit: '厘米',
                        min: 0.5,
                        max: 15,
                        testId: 'page-margin-top',
                      }),
                      numInput('下边距', marginBottomCm.value, (n) => (marginBottomCm.value = n), {
                        unit: '厘米',
                        min: 0.5,
                        max: 15,
                        testId: 'page-margin-bottom',
                      }),
                      numInput('左边距', marginLeftCm.value, (n) => (marginLeftCm.value = n), {
                        unit: '厘米',
                        min: 0.5,
                        max: 15,
                        testId: 'page-margin-left',
                      }),
                      numInput('右边距', marginRightCm.value, (n) => (marginRightCm.value = n), {
                        unit: '厘米',
                        min: 0.5,
                        max: 15,
                        testId: 'page-margin-right',
                      }),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '应用于'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            value: pageScope.value,
                            onChange: (e: Event) => {
                              pageScope.value = (e.target as HTMLSelectElement).value as
                                'document' | 'section';
                            },
                          },
                          [
                            h('option', { value: 'document' }, '整篇文档'),
                            h('option', { value: 'section' }, '本节'),
                          ]
                        ),
                      ]),
                    ]),
                  ]),
              errorText.value
                ? h('div', { class: 'weboffice-dialog__error', role: 'alert' }, errorText.value)
                : null,
              h('footer', { class: 'weboffice-dialog__foot' }, [
                h(
                  'button',
                  {
                    type: 'button',
                    class: 'weboffice-dialog__btn',
                    onClick: handleClose,
                  },
                  '取消'
                ),
                h(
                  'button',
                  {
                    type: 'submit',
                    class: 'weboffice-dialog__btn weboffice-dialog__btn--primary',
                    'data-testid': 'official-format-apply',
                  },
                  '确定应用'
                ),
              ]),
            ]),
          ]
        ),
      ]);
    };
  },
});
