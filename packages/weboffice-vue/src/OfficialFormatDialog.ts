import { Teleport, defineComponent, h, nextTick, reactive, ref, watch, type PropType } from 'vue';
import {
  useDocxEditor,
  useEditorState,
  usePageSetup,
  useParagraphFormat,
} from '@docx-editor.dev/vue';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import type { PageSetupUpdate, ParagraphFormatUpdate } from '@docx-editor.dev/vue';
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

type AlignDraft = '' | 'left' | 'center' | 'right' | 'justify';
type SpecialKindDraft = '' | 'none' | 'firstLine' | 'hanging';
type LineRuleDraft = '' | 'exact' | 'atLeast' | 'multiple';

/**
 * 中文公文排版设置对话框（厘米 / 字符 / 磅）：
 * - 段落排版页签：逐字段脏值追踪，支持多段混选保留未触碰属性；首行/悬挂缩进在“字符”单位下写入 OOXML `*Chars` 语义；
 * - 页面设置页签：完整保留自定义纸张尺寸与分节差异，仅提交已改动字段；所有数值框采用 `step="any"` 兼容 2.54 / 3.18 cm 等预设；
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

    // 段落草稿态（字符串存储数值框以支持混选留空显示）
    const align = ref<AlignDraft>('left');
    const indentLeftStr = ref('0');
    const indentRightStr = ref('0');
    const specialKind = ref<SpecialKindDraft>('none');
    const specialUnit = ref<'chars' | 'cm'>('chars');
    const specialValStr = ref('2');
    const spaceBeforeStr = ref('0');
    const spaceAfterStr = ref('0');
    const lineRule = ref<LineRuleDraft>('exact');
    const lineValStr = ref('28');

    const paraMixed = reactive({
      align: false,
      indentLeft: false,
      indentRight: false,
      special: false,
      spaceBefore: false,
      spaceAfter: false,
      lineSpacing: false,
    });

    const paraDirty = reactive({
      align: false,
      indentLeft: false,
      indentRight: false,
      special: false,
      spaceBefore: false,
      spaceAfter: false,
      lineSpacing: false,
    });

    // 页面草稿态
    const paperId = ref('a4');
    const customPaperSize = ref<{ widthTwips: number; heightTwips: number } | null>(null);
    const orientation = ref<'portrait' | 'landscape'>('portrait');
    const marginTopStr = ref(String(GB_T_9704_MARGINS_CM.top));
    const marginBottomStr = ref(String(GB_T_9704_MARGINS_CM.bottom));
    const marginLeftStr = ref(String(GB_T_9704_MARGINS_CM.left));
    const marginRightStr = ref(String(GB_T_9704_MARGINS_CM.right));
    const pageScope = ref<'document' | 'section'>('document');

    const pageDirty = reactive({
      paper: false,
      orientation: false,
      marginTop: false,
      marginBottom: false,
      marginLeft: false,
      marginRight: false,
    });

    function resetDirty() {
      paraDirty.align = false;
      paraDirty.indentLeft = false;
      paraDirty.indentRight = false;
      paraDirty.special = false;
      paraDirty.spaceBefore = false;
      paraDirty.spaceAfter = false;
      paraDirty.lineSpacing = false;

      pageDirty.paper = false;
      pageDirty.orientation = false;
      pageDirty.marginTop = false;
      pageDirty.marginBottom = false;
      pageDirty.marginLeft = false;
      pageDirty.marginRight = false;
    }

    function seedFromEditor() {
      errorText.value = '';
      resetDirty();

      // Seed paragraph fields, respecting mixed-selection disagreements
      const pf = paragraph.format.value;
      const fontPt = currentFontSizePt.value;
      if (pf) {
        paraMixed.align = pf.disagrees.alignment;
        align.value = pf.disagrees.alignment ? '' : (pf.alignment ?? 'left');

        paraMixed.indentLeft = pf.disagrees.indentLeft || pf.indentUnknown;
        indentLeftStr.value = paraMixed.indentLeft
          ? ''
          : String(twipsToCm(pf.indentLeftTwips ?? 0));

        paraMixed.indentRight = pf.disagrees.indentRight || pf.indentUnknown;
        indentRightStr.value = paraMixed.indentRight
          ? ''
          : String(twipsToCm(pf.indentRightTwips ?? 0));

        paraMixed.special = pf.disagrees.indentFirstLine || pf.indentUnknown;
        if (paraMixed.special) {
          specialKind.value = '';
          specialUnit.value = 'chars';
          specialValStr.value = '';
        } else {
          const fl = pf.indentFirstLineTwips ?? 0;
          if (fl > 0) {
            specialKind.value = 'firstLine';
            specialUnit.value = 'chars';
            specialValStr.value = String(twipsToChars(fl, fontPt));
          } else if (fl < 0) {
            specialKind.value = 'hanging';
            specialUnit.value = 'cm';
            specialValStr.value = String(twipsToCm(Math.abs(fl)));
          } else {
            specialKind.value = 'none';
            specialUnit.value = 'chars';
            specialValStr.value = '2';
          }
        }

        paraMixed.spaceBefore = pf.disagrees.spaceBeforePt;
        spaceBeforeStr.value = paraMixed.spaceBefore
          ? ''
          : String(Math.round((pf.spaceBeforePt ?? 0) * 10) / 10);

        paraMixed.spaceAfter = pf.disagrees.spaceAfterPt;
        spaceAfterStr.value = paraMixed.spaceAfter
          ? ''
          : String(Math.round((pf.spaceAfterPt ?? 0) * 10) / 10);

        paraMixed.lineSpacing = pf.disagrees.lineSpacing;
        if (paraMixed.lineSpacing) {
          lineRule.value = '';
          lineValStr.value = '';
        } else if (pf.lineSpacing) {
          lineRule.value = pf.lineSpacing.rule;
          lineValStr.value = String(Math.round(pf.lineSpacing.value * 100) / 100);
        } else {
          lineRule.value = 'multiple';
          lineValStr.value = '1';
        }
      }

      // Seed page setup fields, preserving custom paper sizes intact
      const ps = pageSetup.pageSetup.value;
      if (ps) {
        orientation.value = ps.orientation;
        const shortSide = Math.min(ps.pageWidthTwips, ps.pageHeightTwips);
        const longSide = Math.max(ps.pageWidthTwips, ps.pageHeightTwips);
        const matched = OFFICIAL_PAPER_SIZES.find(
          (p) => Math.abs(p.widthTwips - shortSide) < 40 && Math.abs(p.heightTwips - longSide) < 40
        );
        if (matched) {
          paperId.value = matched.id;
          customPaperSize.value = null;
        } else {
          paperId.value = 'custom';
          customPaperSize.value = { widthTwips: shortSide, heightTwips: longSide };
        }
        marginTopStr.value = String(twipsToCm(ps.marginsTwips.top));
        marginBottomStr.value = String(twipsToCm(ps.marginsTwips.bottom));
        marginLeftStr.value = String(twipsToCm(ps.marginsTwips.left));
        marginRightStr.value = String(twipsToCm(ps.marginsTwips.right));
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
      const parsed = Number(specialValStr.value);
      if (Number.isFinite(parsed)) {
        const currentTwips =
          specialUnit.value === 'chars' ? charsToTwips(parsed, fontPt) : cmToTwips(parsed);
        specialValStr.value = String(
          nextUnit === 'chars' ? twipsToChars(currentTwips, fontPt) : twipsToCm(currentTwips)
        );
      }
      specialUnit.value = nextUnit;
      paraDirty.special = true;
      paraMixed.special = false;
    }

    function applyOfficialBodyPreset() {
      align.value = 'justify';
      indentLeftStr.value = '0';
      indentRightStr.value = '0';
      specialKind.value = 'firstLine';
      specialUnit.value = 'chars';
      specialValStr.value = '2';
      spaceBeforeStr.value = '0';
      spaceAfterStr.value = '0';
      lineRule.value = 'exact';
      lineValStr.value = '28';
      paraMixed.align = false;
      paraMixed.indentLeft = false;
      paraMixed.indentRight = false;
      paraMixed.special = false;
      paraMixed.spaceBefore = false;
      paraMixed.spaceAfter = false;
      paraMixed.lineSpacing = false;
      paraDirty.align = true;
      paraDirty.indentLeft = true;
      paraDirty.indentRight = true;
      paraDirty.special = true;
      paraDirty.spaceBefore = true;
      paraDirty.spaceAfter = true;
      paraDirty.lineSpacing = true;
    }

    function applyOfficialMarginPreset() {
      paperId.value = 'a4';
      orientation.value = 'portrait';
      marginTopStr.value = String(GB_T_9704_MARGINS_CM.top);
      marginBottomStr.value = String(GB_T_9704_MARGINS_CM.bottom);
      marginLeftStr.value = String(GB_T_9704_MARGINS_CM.left);
      marginRightStr.value = String(GB_T_9704_MARGINS_CM.right);
      pageDirty.paper = true;
      pageDirty.orientation = true;
      pageDirty.marginTop = true;
      pageDirty.marginBottom = true;
      pageDirty.marginLeft = true;
      pageDirty.marginRight = true;
    }

    function applyStandardMarginPreset() {
      // 常规预设只改四边页边距，不覆盖当前文档的自定义或既有纸张尺寸
      marginTopStr.value = '2.54';
      marginBottomStr.value = '2.54';
      marginLeftStr.value = '3.18';
      marginRightStr.value = '3.18';
      pageDirty.marginTop = true;
      pageDirty.marginBottom = true;
      pageDirty.marginLeft = true;
      pageDirty.marginRight = true;
    }

    function handleClose() {
      if (dialogRef.value?.open) dialogRef.value.close();
      props.onClose();
      editorRef.value?.focus();
    }

    function parseNonNegative(raw: string, fallback = 0): number | null {
      const trimmed = raw.trim();
      if (trimmed === '') return null;
      const n = Number(trimmed);
      if (!Number.isFinite(n) || n < 0) return fallback;
      return n;
    }

    function handleSubmit(e: Event) {
      e.preventDefault();
      errorText.value = '';

      if (activeTab.value === 'paragraph') {
        const update: { -readonly [K in keyof ParagraphFormatUpdate]: ParagraphFormatUpdate[K] } =
          {};
        let hasChanges = false;

        if (paraDirty.align && align.value !== '') {
          update.alignment = align.value;
          hasChanges = true;
        }
        if (paraDirty.indentLeft) {
          const val = parseNonNegative(indentLeftStr.value);
          if (val !== null) {
            update.indentLeftTwips = cmToTwips(Math.min(50, val));
            hasChanges = true;
          }
        }
        if (paraDirty.indentRight) {
          const val = parseNonNegative(indentRightStr.value);
          if (val !== null) {
            update.indentRightTwips = cmToTwips(Math.min(50, val));
            hasChanges = true;
          }
        }
        if (paraDirty.special && specialKind.value !== '') {
          const fontPt = currentFontSizePt.value;
          if (specialKind.value === 'none') {
            update.indentFirstLineChars = 0;
            update.indentFirstLineTwips = 0;
            hasChanges = true;
          } else {
            const rawVal = parseNonNegative(specialValStr.value) ?? 0;
            const sign = specialKind.value === 'hanging' ? -1 : 1;
            if (specialUnit.value === 'chars') {
              const charsHundredths = Math.round(Math.min(50, rawVal) * 100);
              update.indentFirstLineChars = sign * charsHundredths;
              update.indentFirstLineTwips = sign * charsToTwips(Math.min(50, rawVal), fontPt);
            } else {
              update.indentFirstLineTwips = sign * cmToTwips(Math.min(50, rawVal));
            }
            hasChanges = true;
          }
        }
        if (paraDirty.spaceBefore) {
          const val = parseNonNegative(spaceBeforeStr.value);
          if (val !== null) {
            update.spaceBeforePt = Math.min(1584, val);
            hasChanges = true;
          }
        }
        if (paraDirty.spaceAfter) {
          const val = parseNonNegative(spaceAfterStr.value);
          if (val !== null) {
            update.spaceAfterPt = Math.min(1584, val);
            hasChanges = true;
          }
        }
        if (paraDirty.lineSpacing && lineRule.value !== '') {
          const rawVal = Number(lineValStr.value);
          if (Number.isFinite(rawVal)) {
            update.lineSpacing = {
              rule: lineRule.value,
              value:
                lineRule.value === 'multiple'
                  ? Math.max(0.5, Math.min(10, rawVal))
                  : Math.max(6, Math.min(200, rawVal)),
            };
            hasChanges = true;
          }
        }

        if (!hasChanges) {
          handleClose();
          return;
        }
        const ok = paragraph.apply(update);
        if (!ok) {
          errorText.value = '当前位置无法修改段落格式（可能为只读或受保护区域）';
          return;
        }
        handleClose();
        return;
      }

      // 页面设置：仅提交被用户改动过的字段，绝不覆写自定义纸张或未改动项
      const pageUpdate: { -readonly [K in keyof PageSetupUpdate]: PageSetupUpdate[K] } = {};
      let hasPageChanges = false;

      if (pageDirty.paper) {
        if (paperId.value === 'custom' && customPaperSize.value) {
          const base = customPaperSize.value;
          pageUpdate.pageWidthTwips =
            orientation.value === 'landscape' ? base.heightTwips : base.widthTwips;
          pageUpdate.pageHeightTwips =
            orientation.value === 'landscape' ? base.widthTwips : base.heightTwips;
          pageUpdate.orientation = orientation.value;
          hasPageChanges = true;
        } else {
          const paper = OFFICIAL_PAPER_SIZES.find((p) => p.id === paperId.value);
          if (paper) {
            pageUpdate.pageWidthTwips =
              orientation.value === 'landscape' ? paper.heightTwips : paper.widthTwips;
            pageUpdate.pageHeightTwips =
              orientation.value === 'landscape' ? paper.widthTwips : paper.heightTwips;
            pageUpdate.orientation = orientation.value;
            hasPageChanges = true;
          }
        }
      } else if (pageDirty.orientation) {
        // 仅改纸张方向时不带固定宽高，让引擎按各节自身尺寸翻转，保护分节自定义纸型
        pageUpdate.orientation = orientation.value;
        hasPageChanges = true;
      }

      const marginChecks: ReadonlyArray<{
        dirty: boolean;
        raw: string;
        key: 'marginTopTwips' | 'marginBottomTwips' | 'marginLeftTwips' | 'marginRightTwips';
      }> = [
        { dirty: pageDirty.marginTop, raw: marginTopStr.value, key: 'marginTopTwips' },
        { dirty: pageDirty.marginBottom, raw: marginBottomStr.value, key: 'marginBottomTwips' },
        { dirty: pageDirty.marginLeft, raw: marginLeftStr.value, key: 'marginLeftTwips' },
        { dirty: pageDirty.marginRight, raw: marginRightStr.value, key: 'marginRightTwips' },
      ];

      for (const item of marginChecks) {
        if (!item.dirty) continue;
        const val = Number(item.raw);
        if (!Number.isFinite(val) || val < 0 || val > 25) {
          errorText.value = '请输入合法的页边距数值（0 ~ 25 厘米）';
          return;
        }
        pageUpdate[item.key] = cmToTwips(val);
        hasPageChanges = true;
      }

      if (!hasPageChanges) {
        handleClose();
        return;
      }

      pageUpdate.scope = pageScope.value;
      const ok = pageSetup.apply(pageUpdate);
      if (!ok) {
        errorText.value = '当前文档无法应用页面设置（请检查页边距是否过大或处于只读模式）';
        return;
      }
      handleClose();
    }

    const numInput = (
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
    ) =>
      h('label', { class: 'weboffice-dialog__field' }, [
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

    return () => {
      if (!props.open || !props.hostElement) return null;
      const customLabel = customPaperSize.value
        ? `自定义 (${twipsToCm(customPaperSize.value.widthTwips)} × ${twipsToCm(customPaperSize.value.heightTwips)} cm)`
        : '自定义纸张';

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
                            'data-testid': 'para-align',
                            value: align.value,
                            onChange: (e: Event) => {
                              align.value = (e.target as HTMLSelectElement).value as AlignDraft;
                              paraMixed.align = false;
                              paraDirty.align = true;
                            },
                          },
                          [
                            paraMixed.align
                              ? h('option', { value: '', disabled: true }, '— 保持混合 —')
                              : null,
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
                      numInput(
                        '左缩进',
                        indentLeftStr.value,
                        (v) => {
                          indentLeftStr.value = v;
                          paraMixed.indentLeft = false;
                          paraDirty.indentLeft = true;
                        },
                        {
                          unit: '厘米',
                          placeholder: paraMixed.indentLeft ? '混合' : undefined,
                        }
                      ),
                      numInput(
                        '右缩进',
                        indentRightStr.value,
                        (v) => {
                          indentRightStr.value = v;
                          paraMixed.indentRight = false;
                          paraDirty.indentRight = true;
                        },
                        {
                          unit: '厘米',
                          placeholder: paraMixed.indentRight ? '混合' : undefined,
                        }
                      ),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '特殊缩进'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            'data-testid': 'para-special-kind',
                            value: specialKind.value,
                            onChange: (e: Event) => {
                              const nextKind = (e.target as HTMLSelectElement)
                                .value as SpecialKindDraft;
                              specialKind.value = nextKind;
                              if (nextKind !== 'none' && specialValStr.value.trim() === '') {
                                specialValStr.value = specialUnit.value === 'chars' ? '2' : '0.74';
                              }
                              paraMixed.special = false;
                              paraDirty.special = true;
                            },
                          },
                          [
                            paraMixed.special
                              ? h('option', { value: '', disabled: true }, '— 保持混合 —')
                              : null,
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
                            inputmode: 'decimal',
                            class: 'weboffice-dialog__input',
                            'data-testid': 'para-special-val',
                            disabled: specialKind.value === 'none' || specialKind.value === '',
                            step: 'any',
                            min: 0,
                            max: 50,
                            placeholder: paraMixed.special ? '混合' : undefined,
                            value: specialValStr.value,
                            onInput: (e: Event) => {
                              specialValStr.value = (e.target as HTMLInputElement).value;
                              paraMixed.special = false;
                              paraDirty.special = true;
                            },
                          }),
                          h(
                            'select',
                            {
                              class: 'weboffice-dialog__unit-select',
                              disabled: specialKind.value === 'none' || specialKind.value === '',
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
                      numInput(
                        '段前间距',
                        spaceBeforeStr.value,
                        (v) => {
                          spaceBeforeStr.value = v;
                          paraMixed.spaceBefore = false;
                          paraDirty.spaceBefore = true;
                        },
                        {
                          unit: '磅',
                          max: 200,
                          placeholder: paraMixed.spaceBefore ? '混合' : undefined,
                        }
                      ),
                      numInput(
                        '段后间距',
                        spaceAfterStr.value,
                        (v) => {
                          spaceAfterStr.value = v;
                          paraMixed.spaceAfter = false;
                          paraDirty.spaceAfter = true;
                        },
                        {
                          unit: '磅',
                          max: 200,
                          placeholder: paraMixed.spaceAfter ? '混合' : undefined,
                        }
                      ),
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
                                .value as LineRuleDraft;
                              if (nextRule === 'multiple' && lineRule.value !== 'multiple') {
                                lineValStr.value = '1.5';
                              } else if (
                                nextRule !== 'multiple' &&
                                (lineRule.value === 'multiple' || lineRule.value === '')
                              ) {
                                lineValStr.value = '28';
                              }
                              lineRule.value = nextRule;
                              paraMixed.lineSpacing = false;
                              paraDirty.lineSpacing = true;
                            },
                          },
                          [
                            paraMixed.lineSpacing
                              ? h('option', { value: '', disabled: true }, '— 保持混合 —')
                              : null,
                            h('option', { value: 'exact' }, '固定值（公文推荐）'),
                            h('option', { value: 'atLeast' }, '最小值'),
                            h('option', { value: 'multiple' }, '多倍行距'),
                          ]
                        ),
                      ]),
                      numInput(
                        '行距设置值',
                        lineValStr.value,
                        (v) => {
                          lineValStr.value = v;
                          if (lineRule.value === '') lineRule.value = 'exact';
                          paraMixed.lineSpacing = false;
                          paraDirty.lineSpacing = true;
                        },
                        {
                          unit: lineRule.value === 'multiple' ? '倍' : '磅',
                          min: lineRule.value === 'multiple' ? 0.5 : 6,
                          max: lineRule.value === 'multiple' ? 10 : 200,
                          testId: 'para-line-val',
                          placeholder: paraMixed.lineSpacing ? '混合' : undefined,
                        }
                      ),
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
                          'data-testid': 'preset-standard-margins',
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
                            'data-testid': 'page-paper-size',
                            value: paperId.value,
                            onChange: (e: Event) => {
                              paperId.value = (e.target as HTMLSelectElement).value;
                              pageDirty.paper = true;
                            },
                          },
                          [
                            ...OFFICIAL_PAPER_SIZES.map((p) =>
                              h('option', { key: p.id, value: p.id }, p.label)
                            ),
                            customPaperSize.value
                              ? h('option', { key: 'custom', value: 'custom' }, customLabel)
                              : null,
                          ]
                        ),
                      ]),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '纸张方向'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            'data-testid': 'page-orientation',
                            value: orientation.value,
                            onChange: (e: Event) => {
                              orientation.value = (e.target as HTMLSelectElement).value as
                                'portrait' | 'landscape';
                              pageDirty.orientation = true;
                            },
                          },
                          [
                            h('option', { value: 'portrait' }, '纵向'),
                            h('option', { value: 'landscape' }, '横向'),
                          ]
                        ),
                      ]),
                      numInput(
                        '上边距',
                        marginTopStr.value,
                        (v) => {
                          marginTopStr.value = v;
                          pageDirty.marginTop = true;
                        },
                        {
                          unit: '厘米',
                          min: 0,
                          max: 25,
                          testId: 'page-margin-top',
                        }
                      ),
                      numInput(
                        '下边距',
                        marginBottomStr.value,
                        (v) => {
                          marginBottomStr.value = v;
                          pageDirty.marginBottom = true;
                        },
                        {
                          unit: '厘米',
                          min: 0,
                          max: 25,
                          testId: 'page-margin-bottom',
                        }
                      ),
                      numInput(
                        '左边距',
                        marginLeftStr.value,
                        (v) => {
                          marginLeftStr.value = v;
                          pageDirty.marginLeft = true;
                        },
                        {
                          unit: '厘米',
                          min: 0,
                          max: 25,
                          testId: 'page-margin-left',
                        }
                      ),
                      numInput(
                        '右边距',
                        marginRightStr.value,
                        (v) => {
                          marginRightStr.value = v;
                          pageDirty.marginRight = true;
                        },
                        {
                          unit: '厘米',
                          min: 0,
                          max: 25,
                          testId: 'page-margin-right',
                        }
                      ),
                      h('label', { class: 'weboffice-dialog__field' }, [
                        h('span', { class: 'weboffice-dialog__label' }, '应用于'),
                        h(
                          'select',
                          {
                            class: 'weboffice-dialog__select',
                            'data-testid': 'page-scope',
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
