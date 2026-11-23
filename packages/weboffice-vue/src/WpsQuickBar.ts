import {
  Teleport,
  computed,
  defineComponent,
  h,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
  type PropType,
} from 'vue';
import {
  useDocxEditor,
  useEditorCommand,
  useEditorEvent,
  useEditorState,
  useHeaderFooterState,
  useParagraphFormat,
} from '@docx-editor.dev/vue';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import type {
  ToggleListPreset,
  Watermark,
  WatermarkLayout,
} from '@docx-editor.dev/core/contracts/types';
import type { ParagraphBorderEdgeWrite } from '@docx-editor.dev/core/store';
import { charsToTwips } from './units';
import type { OfficialFormatTab } from './OfficialFormatDialog';
import { executeClipboardCopyOrCut, executeClipboardPaste } from './qbar-clipboard';
import type { UnderlineStylePreset } from './p1-presets';
import { buildCustomWatermark } from './p2-presets';
import {
  clampPopoverWithinViewport,
  renderBordersMenu,
  renderNumberingMenu,
  renderTableContextSection,
  renderTextEffectsMenu,
} from './qbar-p1-menus';
import {
  renderLineSpacingMenu,
  renderPageExtrasMenu,
  renderPasteMenu,
  renderSymbolsMenu,
  renderTableAdvancedSection,
  renderTableInsertGrid,
  type HeaderFooterQuickAction,
} from './qbar-p2-menus';
import {
  insertRasterImageFromFile,
  runHeaderFooterQuickAction,
  runTableQuickAction,
  type TableContextQuickAction,
} from './qbar-p2-actions';

type QuickPopover =
  | 'none'
  | 'lineSpacing'
  | 'numbering'
  | 'borders'
  | 'textEffects'
  | 'pageExtras'
  | 'table'
  | 'symbols'
  | 'paste';

const selectFontSizePt = (snapshot: EditorSnapshot): number => {
  const pt = snapshot.formatting?.fontSizePt;
  return typeof pt === 'number' && pt > 0 ? pt : 16;
};

const selectIsDisabled = (snapshot: EditorSnapshot): boolean =>
  snapshot.isLoading || snapshot.isOpening === true || !snapshot.editable;

const selectInTable = (snapshot: EditorSnapshot): boolean => snapshot.table !== null;
const selectTableHeaderCount = (snapshot: EditorSnapshot): number =>
  snapshot.table?.headerRowCount ?? 0;
const selectIsTableHeaderRow = (snapshot: EditorSnapshot): boolean =>
  Boolean(snapshot.table?.isHeaderRow);
const selectCanToggleHeaderRow = (snapshot: EditorSnapshot): boolean =>
  snapshot.table !== null && snapshot.table.rowIndex <= (snapshot.table.headerRowCount ?? 0);

/**
 * WPS 风格公文快捷增强工具栏（挂载于主工具栏下方，补齐 WPS 顶部栏高频 P0 + P1 能力）：
 * - P0：剪贴板（富文本复制/剪切/粘贴/无格式粘贴）、平铺对齐与段落标记 ¶、公文行距/首行2字符/厘米页面设置、表格网格/本地图片/公文符号、查找替换
 * - P1：GB/T 9704 中文公文编号（一、/（一）/1./（1））、段落边框（含公文红线/双红线）与底纹、下划线线型与字符间距、分页/分节符直达、表格上下文操作
 */
export const WpsQuickBar = defineComponent({
  name: 'WpsQuickBar',
  props: {
    hostElement: {
      type: Object as PropType<HTMLElement | null>,
      default: null,
    },
    findReplaceOpen: {
      type: Boolean,
      default: false,
    },
    onToggleFindReplace: {
      type: Function as PropType<() => void>,
      required: true,
    },
    onOpenOfficialDialog: {
      type: Function as PropType<(tab: OfficialFormatTab) => void>,
      required: true,
    },
  },
  setup(props) {
    // oxlint-disable-next-line react/rules-of-hooks
    const editorRef = useDocxEditor();
    // oxlint-disable-next-line react/rules-of-hooks
    const alignLeft = useEditorCommand('alignment.left');
    // oxlint-disable-next-line react/rules-of-hooks
    const alignCenter = useEditorCommand('alignment.center');
    // oxlint-disable-next-line react/rules-of-hooks
    const alignRight = useEditorCommand('alignment.right');
    // oxlint-disable-next-line react/rules-of-hooks
    const alignJustify = useEditorCommand('alignment.justify');
    // oxlint-disable-next-line react/rules-of-hooks
    const paraMarks = useEditorCommand('review.paragraphMarks');
    // oxlint-disable-next-line react/rules-of-hooks
    const tableInsertCmd = useEditorCommand('table.insert');
    // oxlint-disable-next-line react/rules-of-hooks
    const tableMergeCmd = useEditorCommand('table.mergeCells');
    // oxlint-disable-next-line react/rules-of-hooks
    const imageInsertCmd = useEditorCommand('image.insert');
    // oxlint-disable-next-line react/rules-of-hooks
    const formatPainterCmd = useEditorCommand('format.painter');
    // oxlint-disable-next-line react/rules-of-hooks
    const clearFormatCmd = useEditorCommand('format.clear');
    // oxlint-disable-next-line react/rules-of-hooks
    const footnoteCmd = useEditorCommand('insert.footnote');
    // oxlint-disable-next-line react/rules-of-hooks
    const endnoteCmd = useEditorCommand('insert.endnote');
    // oxlint-disable-next-line react/rules-of-hooks
    const textBoxCmd = useEditorCommand('insert.textBox');
    // oxlint-disable-next-line react/rules-of-hooks
    const hfState = useHeaderFooterState();
    // oxlint-disable-next-line react/rules-of-hooks
    const paragraph = useParagraphFormat();
    // oxlint-disable-next-line react/rules-of-hooks
    const currentFontSizePt = useEditorState(selectFontSizePt);
    // oxlint-disable-next-line react/rules-of-hooks
    const isDisabled = useEditorState(selectIsDisabled);
    // oxlint-disable-next-line react/rules-of-hooks
    const inTable = useEditorState(selectInTable);
    // oxlint-disable-next-line react/rules-of-hooks
    const tableHeaderRowCount = useEditorState(selectTableHeaderCount);
    // oxlint-disable-next-line react/rules-of-hooks
    const isTableHeaderRow = useEditorState(selectIsTableHeaderRow);
    // oxlint-disable-next-line react/rules-of-hooks
    const canToggleHeaderRow = useEditorState(selectCanToggleHeaderRow);

    const anchorRef = ref<HTMLElement | null>(null);
    const chromeBandEl = shallowRef<HTMLElement | null>(null);
    const barRef = ref<HTMLElement | null>(null);
    const fileInputRef = ref<HTMLInputElement | null>(null);
    const openPopover = ref<QuickPopover>('none');
    const tableHover = ref({ rows: 3, cols: 4 });
    const splitLimits = ref<{ maxRows: number; maxCols: number } | null>(null);
    const splitRows = ref(1);
    const splitCols = ref(2);
    const customSpacingPt = ref('1.5');
    const watermarkState = ref<{ has: boolean; text: string }>({ has: false, text: '' });
    const customWmText = ref('');
    const customWmLayout = ref<WatermarkLayout>('diagonal');
    const customWmColor = ref('C0C0C0');
    const localTitlePage = ref(false);
    const toastMsg = ref('');
    const imageInsertBusy = ref(false);
    let cutSeq = 0;
    let toastTimer: number | null = null;
    let observer: MutationObserver | undefined;

    const hasFirstLineIndent = computed(() => {
      const pf = paragraph.format.value;
      if (!pf || pf.indentUnknown || pf.disagrees.indentFirstLine) return false;
      const fl = pf.indentFirstLineTwips;
      if (fl === null || fl <= 0) return false;
      const twoCharsTwips = charsToTwips(2, currentFontSizePt.value);
      return Math.abs(fl - twoCharsTwips) <= 20;
    });

    function showToast(message: string) {
      toastMsg.value = message;
      if (toastTimer !== null) window.clearTimeout(toastTimer);
      toastTimer = window.setTimeout(() => {
        toastMsg.value = '';
        toastTimer = null;
      }, 2600);
    }

    function locateChromeBand() {
      const host = props.hostElement ?? anchorRef.value?.closest<HTMLElement>('.weboffice-host');
      if (!host) return;
      const toolbar = host.querySelector<HTMLElement>('.docx-toolbar');
      const band = toolbar?.parentElement?.parentElement ?? toolbar?.parentElement ?? null;
      chromeBandEl.value = band;
    }

    function onDocPointerDown(event: PointerEvent) {
      if (openPopover.value === 'none') return;
      const target = event.target instanceof Node ? event.target : null;
      if (target && barRef.value?.contains(target)) return;
      openPopover.value = 'none';
    }

    function clampActivePopover() {
      if (openPopover.value === 'none') return;
      clampPopoverWithinViewport(barRef.value, props.hostElement);
    }

    function syncDynamicPopoverState(next: QuickPopover) {
      const ed = editorRef.value;
      if (!ed) return;
      if (next === 'pageExtras') {
        const sf = ed.query({ type: 'sectionFurniture' });
        localTitlePage.value = Boolean(hfState.value?.titlePage ?? sf?.titlePage ?? false);
        const wm = ed.query({ type: 'watermark' });
        watermarkState.value = { has: wm !== null, text: wm?.text ?? '' };
        if (wm?.text) customWmText.value = wm.text;
        if (wm?.layout) customWmLayout.value = wm.layout;
        if (wm?.color) customWmColor.value = wm.color.toUpperCase();
      } else if (next === 'table' && inTable.value) {
        const cfg = ed.query({ type: 'splitCellConfig' });
        splitLimits.value = cfg;
        if (cfg) {
          splitRows.value = Math.min(Math.max(1, splitRows.value), cfg.maxRows);
          splitCols.value = Math.min(Math.max(1, splitCols.value), cfg.maxCols);
        }
      }
    }

    function resetDocLocalState() {
      localTitlePage.value = false;
      watermarkState.value = { has: false, text: '' };
    }

    // oxlint-disable-next-line react/rules-of-hooks
    useEditorEvent('change', (ev) => {
      if (ev.source === 'load') resetDocLocalState();
    });

    watch(editorRef, () => {
      resetDocLocalState();
    });

    watch(openPopover, async (next) => {
      if (next === 'none') return;
      syncDynamicPopoverState(next);
      await nextTick();
      clampActivePopover();
    });

    onMounted(() => {
      document.addEventListener('pointerdown', onDocPointerDown, true);
      window.addEventListener('resize', clampActivePopover);
      locateChromeBand();
      const host = props.hostElement ?? anchorRef.value?.closest<HTMLElement>('.weboffice-host');
      if (host) {
        observer = new MutationObserver(locateChromeBand);
        observer.observe(host, { childList: true, subtree: true });
      }
    });

    onBeforeUnmount(() => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      window.removeEventListener('resize', clampActivePopover);
      observer?.disconnect();
      if (toastTimer !== null) window.clearTimeout(toastTimer);
    });

    function togglePopover(next: QuickPopover) {
      openPopover.value = openPopover.value === next ? 'none' : next;
    }

    async function runCopyOrCut(action: 'copy' | 'cut') {
      const ed = editorRef.value;
      if (!ed) return;
      const cutTicket = action === 'cut' ? ++cutSeq : 0;
      const msg = await executeClipboardCopyOrCut({
        ed,
        getCurrentEditor: () => editorRef.value,
        action,
        cutTicket,
        getLatestCutTicket: () => cutSeq,
      });
      showToast(msg);
    }

    async function runPaste(plainOnly: boolean) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const err = await executeClipboardPaste(ed, plainOnly);
      if (err) showToast(err);
    }

    function applyLineSpacingPreset(rule: 'exact' | 'multiple', value: number) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setLineSpacing', rule, value });
      if (!res.ok) showToast('无法设置行距');
      ed.focus();
    }

    function toggleTwoCharFirstLine() {
      const ed = editorRef.value;
      if (!ed) return;
      const turningOff = hasFirstLineIndent.value;
      const targetChars = turningOff ? 0 : 200;
      const targetTwips = turningOff ? 0 : charsToTwips(2, currentFontSizePt.value);
      const res = ed.exec({
        type: 'setParagraphFormat',
        indentFirstLineChars: targetChars,
        indentFirstLineTwips: targetTwips,
      });
      if (!res.ok) {
        showToast('当前位置无法调整首行缩进');
      } else {
        showToast(turningOff ? '已取消首行缩进' : '已设首行缩进 2 字符');
      }
      ed.focus();
    }

    function applyNumberingPreset(preset: ToggleListPreset) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'toggleList', kind: 'ordered', preset });
      if (!res.ok) {
        showToast('当前位置无法设置编号');
      } else {
        showToast(preset === 'clear' ? '已取消编号' : '已应用公文编号');
      }
      ed.focus();
    }

    function restartNumbering() {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'toggleList', kind: 'ordered', restart: true });
      if (!res.ok) {
        showToast('当前位置无法重新编号');
      } else {
        showToast('已从新序列开始编号');
      }
      ed.focus();
    }

    function applyParagraphBorder(
      edges: readonly ParagraphBorderEdgeWrite[],
      label: string,
      mode: 'merge' | 'replace'
    ) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setParagraphFormat', borders: edges, borderMode: mode });
      if (!res.ok) {
        showToast('当前位置无法设置段落边框');
      } else {
        showToast(`已应用${label}`);
      }
      ed.focus();
    }

    function applyParagraphShading(hex: string | null, label: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setParagraphFormat', shading: hex });
      if (!res.ok) {
        showToast('当前位置无法设置段落底纹');
      } else {
        showToast(hex === null ? '已清除段落底纹' : `已设底纹：${label}`);
      }
      ed.focus();
    }

    function applyUnderlineVariant(value: UnderlineStylePreset['value'], label: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setMarkAttr', mark: 'underline', attr: 'val', value });
      if (!res.ok) {
        showToast('当前位置无法设置下划线样式');
      } else {
        showToast(value === 'none' ? '已取消下划线' : `已设${label}`);
      }
      ed.focus();
    }

    function applyCharSpacing(twips: number, label: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({
        type: 'setMarkAttr',
        mark: 'characterSpacing',
        attr: 'val',
        value: twips,
      });
      if (!res.ok) {
        showToast('当前位置无法设置字符间距');
      } else {
        showToast(`已设字符间距：${label}`);
      }
      ed.focus();
    }

    function applyCustomCharSpacing() {
      const pt = Number(customSpacingPt.value);
      if (!Number.isFinite(pt) || pt < -20 || pt > 100) {
        showToast('请输入 -20 ~ 100 磅之间的字符间距');
        return;
      }
      applyCharSpacing(Math.round(pt * 20), `${pt} 磅`);
    }

    function runTableContextAction(action: TableContextQuickAction) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const out = runTableQuickAction(ed, action);
      if (out.message) showToast(out.message);
      ed.focus();
    }

    function runSplitCell() {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'splitCell', rows: splitRows.value, cols: splitCols.value });
      showToast(
        res.ok
          ? `已拆分为 ${splitRows.value} 行 × ${splitCols.value} 列`
          : res.reason || '当前单元格无法拆分'
      );
      ed.focus();
    }

    function runTableVerticalAlign(alignment: 'top' | 'center' | 'bottom') {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setTableCellVerticalAlignment', alignment });
      const label = alignment === 'top' ? '顶对齐' : alignment === 'center' ? '垂直居中' : '底对齐';
      showToast(res.ok ? `已设单元格${label}` : res.reason || '当前无法设置单元格对齐');
      ed.focus();
    }

    function runTableCellFill(hex: string | null, label: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({
        type: 'setCellFill',
        color: hex ? { kind: 'hex', value: hex } : null,
      });
      showToast(
        res.ok
          ? hex
            ? `已设单元格底纹：${label}`
            : '已清除单元格底纹'
          : res.reason || '无法设置单元格底纹'
      );
      ed.focus();
    }

    function runWholeTableBorder(scope: 'threeLine' | 'tableAll' | 'tableNone') {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setTableBorders', scope });
      const label =
        scope === 'threeLine' ? '公文三线表' : scope === 'tableAll' ? '整表全框线' : '整表无框线';
      showToast(res.ok ? `已应用${label}` : res.reason || '当前表格无法应用框线样式');
      ed.focus();
    }

    function runHfAction(action: HeaderFooterQuickAction) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const out = runHeaderFooterQuickAction(ed, action);
      const sf = ed.query({ type: 'sectionFurniture' });
      localTitlePage.value = Boolean(
        ed.getHeaderFooterState()?.titlePage ?? sf?.titlePage ?? false
      );
      showToast(out.message);
      ed.focus();
    }

    function applyWatermark(watermark: Watermark | null, label: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setWatermark', watermark });
      if (!res.ok) {
        showToast(res.reason || '当前文档无法设置水印');
      } else {
        watermarkState.value = watermark
          ? { has: true, text: watermark.text ?? '' }
          : { has: false, text: '' };
        showToast(watermark ? `已设置水印：${label}` : '已清除水印');
      }
      ed.focus();
    }

    function applyCustomWatermark() {
      const built = buildCustomWatermark({
        text: customWmText.value,
        layout: customWmLayout.value,
        color: customWmColor.value,
      });
      if (!built.ok) {
        showToast(built.error);
        return;
      }
      applyWatermark(built.watermark, built.watermark.text ?? '自定义水印');
    }

    function insertTable(rows: number, cols: number) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'insertTable', rows, cols });
      if (!res.ok) showToast('当前位置无法插入表格');
      ed.focus();
    }

    function insertBreakKind(kind: 'page' | 'section') {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'insertBreak', kind });
      if (!res.ok) {
        showToast(res.reason || '当前位置无法插入分隔符');
      } else {
        showToast(kind === 'page' ? '已插入分页符' : '已插入下一页分节符');
      }
      ed.focus();
    }

    function insertExtraObject(kind: 'footnote' | 'endnote' | 'textBox') {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res =
        kind === 'textBox'
          ? ed.exec({ type: 'insertTextBox' })
          : ed.exec({ type: 'insertNote', noteKind: kind });
      const label = kind === 'footnote' ? '脚注' : kind === 'endnote' ? '尾注' : '文本框';
      showToast(res.ok ? `已插入${label}` : res.reason || `当前位置无法插入${label}`);
      ed.focus();
    }

    async function handleImagePicked(e: Event) {
      const input = e.target as HTMLInputElement;
      const file = input.files?.[0];
      input.value = '';
      if (!file) return;
      const ed = editorRef.value;
      if (!ed) return;
      if (imageInsertBusy.value) {
        showToast('正在插入图片，请稍候');
        return;
      }
      imageInsertBusy.value = true;
      try {
        const out = await insertRasterImageFromFile(ed, file);
        showToast(out.message);
      } finally {
        imageInsertBusy.value = false;
      }
      ed.focus();
    }

    function insertSymbol(char: string) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'replaceText', text: char });
      if (!res.ok) showToast('当前位置无法插入符号');
      ed.focus();
    }

    const keepFocus = (e: MouseEvent) => e.preventDefault();

    const qBtn = (opts: {
      label: string;
      title: string;
      testId?: string;
      active?: boolean;
      disabled?: boolean;
      accent?: boolean;
      onClick: () => void;
    }) =>
      h(
        'button',
        {
          type: 'button',
          class: [
            'weboffice-qbar__btn',
            opts.active ? 'is-active' : '',
            opts.accent ? 'is-accent' : '',
          ],
          'data-testid': opts.testId,
          title: opts.title,
          disabled: opts.disabled,
          onMousedown: keepFocus,
          onClick: opts.onClick,
        },
        opts.label
      );

    return () =>
      h('span', { ref: anchorRef, style: 'display:none' }, [
        chromeBandEl.value
          ? h(Teleport, { to: chromeBandEl.value }, [
              h(
                'div',
                {
                  ref: barRef,
                  class: 'weboffice-qbar',
                  'data-testid': 'weboffice-wps-bar',
                  role: 'toolbar',
                  'aria-label': 'WPS 公文快捷工具栏',
                },
                [
                  // 1. 剪贴板组
                  h('div', { class: 'weboffice-qbar__group' }, [
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '粘贴 ▾',
                        title: '粘贴 / 仅保留文本',
                        testId: 'qbar-paste-menu',
                        active: openPopover.value === 'paste',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('paste'),
                      }),
                      openPopover.value === 'paste'
                        ? renderPasteMenu((plain) => void runPaste(plain))
                        : null,
                    ]),
                    qBtn({
                      label: '复制',
                      title: '复制选中内容 (Ctrl+C)',
                      testId: 'qbar-copy',
                      onClick: () => void runCopyOrCut('copy'),
                    }),
                    qBtn({
                      label: '剪切',
                      title: '剪切选中内容 (Ctrl+X)',
                      testId: 'qbar-cut',
                      disabled: isDisabled.value,
                      onClick: () => void runCopyOrCut('cut'),
                    }),
                    h(
                      'button',
                      {
                        type: 'button',
                        class: [
                          'weboffice-qbar__btn',
                          formatPainterCmd.isActive.value ? 'is-active' : '',
                          formatPainterCmd.value.value === 'locked' ? 'is-locked' : '',
                        ],
                        'data-testid': 'qbar-format-painter',
                        'data-painter-mode': String(formatPainterCmd.value.value ?? 'off'),
                        'aria-pressed': String(formatPainterCmd.isActive.value),
                        title:
                          formatPainterCmd.value.value === 'locked'
                            ? '格式刷已锁定（连续刷格式，按 Esc 或再次点击退出）'
                            : '格式刷（单击刷一次，双击连续刷）',
                        disabled: !formatPainterCmd.isEnabled.value,
                        onMousedown: keepFocus,
                        onClick: () => formatPainterCmd.execute(),
                      },
                      formatPainterCmd.value.value === 'locked' ? '格式刷🔒' : '格式刷'
                    ),
                    qBtn({
                      label: '清格式',
                      title: '清除选中文字与段落的直接格式',
                      testId: 'qbar-clear-format',
                      disabled: !clearFormatCmd.isEnabled.value,
                      onClick: () => {
                        clearFormatCmd.execute();
                        showToast('已清除直接格式');
                      },
                    }),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 2. 平铺对齐 + 编辑标记 ¶
                  h('div', { class: 'weboffice-qbar__group' }, [
                    qBtn({
                      label: '左',
                      title: '左对齐',
                      testId: 'qbar-align-left',
                      active: alignLeft.isActive.value,
                      disabled: !alignLeft.isEnabled.value,
                      onClick: () => alignLeft.execute(),
                    }),
                    qBtn({
                      label: '中',
                      title: '居中对齐',
                      testId: 'qbar-align-center',
                      active: alignCenter.isActive.value,
                      disabled: !alignCenter.isEnabled.value,
                      onClick: () => alignCenter.execute(),
                    }),
                    qBtn({
                      label: '右',
                      title: '右对齐',
                      testId: 'qbar-align-right',
                      active: alignRight.isActive.value,
                      disabled: !alignRight.isEnabled.value,
                      onClick: () => alignRight.execute(),
                    }),
                    qBtn({
                      label: '两端',
                      title: '两端对齐',
                      testId: 'qbar-align-justify',
                      active: alignJustify.isActive.value,
                      disabled: !alignJustify.isEnabled.value,
                      onClick: () => alignJustify.execute(),
                    }),
                    qBtn({
                      label: '¶',
                      title: '显示/隐藏段落编辑标记',
                      testId: 'qbar-para-marks',
                      active: paraMarks.isActive.value,
                      disabled: !paraMarks.isEnabled.value,
                      onClick: () => paraMarks.execute(),
                    }),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 3. 公文快捷排版组（P0 + P1）
                  h('div', { class: 'weboffice-qbar__group' }, [
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '行距 ▾',
                        title: '公文固定值行距与倍数行距',
                        testId: 'qbar-line-spacing',
                        active: openPopover.value === 'lineSpacing',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('lineSpacing'),
                      }),
                      openPopover.value === 'lineSpacing'
                        ? renderLineSpacingMenu({
                            onApply: applyLineSpacingPreset,
                            onOpenDialog: () => {
                              openPopover.value = 'none';
                              props.onOpenOfficialDialog('paragraph');
                            },
                          })
                        : null,
                    ]),
                    qBtn({
                      label: '首行',
                      title: '一键切换当前段落首行缩进 2 字符',
                      testId: 'qbar-first-line-2char',
                      active: hasFirstLineIndent.value,
                      disabled: isDisabled.value,
                      onClick: toggleTwoCharFirstLine,
                    }),
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '编号 ▾',
                        title: '中文公文编号样式（一、/（一）/1./（1））',
                        testId: 'qbar-cn-numbering',
                        active: openPopover.value === 'numbering',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('numbering'),
                      }),
                      openPopover.value === 'numbering'
                        ? renderNumberingMenu(applyNumberingPreset, restartNumbering)
                        : null,
                    ]),
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '框线 ▾',
                        title: '段落边框（含公文红线/双红线）与段落底纹',
                        testId: 'qbar-para-border',
                        active: openPopover.value === 'borders',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('borders'),
                      }),
                      openPopover.value === 'borders'
                        ? renderBordersMenu({
                            onApplyBorder: applyParagraphBorder,
                            onApplyShading: applyParagraphShading,
                          })
                        : null,
                    ]),
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '字效 ▾',
                        title: '下划线线型与字符间距调节',
                        testId: 'qbar-text-effects',
                        active: openPopover.value === 'textEffects',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('textEffects'),
                      }),
                      openPopover.value === 'textEffects'
                        ? renderTextEffectsMenu({
                            customSpacingPt: customSpacingPt.value,
                            onUpdateCustomPt: (val) => {
                              customSpacingPt.value = val;
                            },
                            onApplyUnderline: applyUnderlineVariant,
                            onApplyCharSpacing: applyCharSpacing,
                            onApplyCustomSpacing: applyCustomCharSpacing,
                          })
                        : null,
                    ]),
                    qBtn({
                      label: '页面',
                      title:
                        '页面设置与分栏（厘米 / GB/T 9704 公文标准页边距 / 多栏 / 页眉页脚边距）',
                      testId: 'qbar-page-setup-cm',
                      disabled: isDisabled.value,
                      onClick: () => {
                        openPopover.value = 'none';
                        props.onOpenOfficialDialog('page');
                      },
                    }),
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: hfState.value?.editing ? '页眉✦▾' : '页眉 ▾',
                        title: '页眉/页脚、公文页码（— 1 —）与公文水印',
                        testId: 'qbar-page-extras',
                        active:
                          openPopover.value === 'pageExtras' || Boolean(hfState.value?.editing),
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('pageExtras'),
                      }),
                      openPopover.value === 'pageExtras'
                        ? renderPageExtrasMenu({
                            editingRegion: hfState.value?.editing ?? null,
                            titlePage: hfState.value?.titlePage ?? localTitlePage.value,
                            hasWatermark: watermarkState.value.has,
                            currentWatermarkText: watermarkState.value.text,
                            customWmText: customWmText.value,
                            customWmLayout: customWmLayout.value,
                            customWmColor: customWmColor.value,
                            onUpdateCustomText: (v) => {
                              customWmText.value = v;
                            },
                            onUpdateCustomLayout: (v) => {
                              customWmLayout.value = v;
                            },
                            onUpdateCustomColor: (v) => {
                              customWmColor.value = v;
                            },
                            onHeaderFooterAction: runHfAction,
                            onApplyWatermarkPreset: applyWatermark,
                            onApplyCustomWatermark: applyCustomWatermark,
                            onClearWatermark: () => applyWatermark(null, ''),
                          })
                        : null,
                    ]),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 4. 快捷插入组（表格含上下文操作 + 本地图片 + 符号与分页/分节符）
                  h('div', { class: 'weboffice-qbar__group' }, [
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: inTable.value ? '表格✦▾' : '表格 ▾',
                        title: inTable.value
                          ? '表格操作（合并单元格/重复表头行/插删行列）与插入新表'
                          : '插入表格',
                        testId: 'qbar-insert-table',
                        active: openPopover.value === 'table' || inTable.value,
                        disabled:
                          isDisabled.value || (!tableInsertCmd.isEnabled.value && !inTable.value),
                        onClick: () => togglePopover('table'),
                      }),
                      openPopover.value === 'table'
                        ? h('div', { class: 'weboffice-qbar__table-grid' }, [
                            inTable.value
                              ? renderTableContextSection({
                                  canMerge: tableMergeCmd.isEnabled.value,
                                  canToggleHeaderRow: canToggleHeaderRow.value,
                                  isHeaderRow: isTableHeaderRow.value,
                                  headerRowCount: tableHeaderRowCount.value,
                                  onAction: runTableContextAction,
                                  advancedSlot: renderTableAdvancedSection({
                                    splitLimits: splitLimits.value,
                                    splitRows: splitRows.value,
                                    splitCols: splitCols.value,
                                    onUpdateSplitRows: (n) => {
                                      splitRows.value = n;
                                    },
                                    onUpdateSplitCols: (n) => {
                                      splitCols.value = n;
                                    },
                                    onSplitCell: runSplitCell,
                                    onVerticalAlign: runTableVerticalAlign,
                                    onCellFill: runTableCellFill,
                                    onWholeTableBorder: runWholeTableBorder,
                                  }),
                                })
                              : null,
                            tableInsertCmd.isEnabled.value
                              ? renderTableInsertGrid({
                                  hover: tableHover.value,
                                  onHover: (rows, cols) => {
                                    tableHover.value = { rows, cols };
                                  },
                                  onInsert: insertTable,
                                })
                              : null,
                          ])
                        : null,
                    ]),
                    qBtn({
                      label: '图片',
                      title: '从本地插入图片 (PNG/JPG/GIF)',
                      testId: 'qbar-insert-image',
                      disabled: !imageInsertCmd.isEnabled.value,
                      onClick: () => {
                        openPopover.value = 'none';
                        fileInputRef.value?.click();
                      },
                    }),
                    h('input', {
                      ref: fileInputRef,
                      type: 'file',
                      accept: 'image/png,image/jpeg,image/gif',
                      style: 'display:none',
                      onChange: handleImagePicked,
                    }),
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '符号 ▾',
                        title: '插入分页符、分节符、脚注/尾注/文本框与公文常用符号',
                        testId: 'qbar-symbols',
                        active: openPopover.value === 'symbols',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('symbols'),
                      }),
                      openPopover.value === 'symbols'
                        ? renderSymbolsMenu({
                            footnoteEnabled: footnoteCmd.isEnabled.value,
                            endnoteEnabled: endnoteCmd.isEnabled.value,
                            textBoxEnabled: textBoxCmd.isEnabled.value,
                            onInsertBreak: insertBreakKind,
                            onInsertExtra: insertExtraObject,
                            onInsertSymbol: insertSymbol,
                          })
                        : null,
                    ]),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 5. 查找替换
                  h('div', { class: 'weboffice-qbar__group' }, [
                    qBtn({
                      label: '查找',
                      title: '查找与替换 (Ctrl+H / Cmd+Option+F)',
                      testId: 'weboffice-open-find-replace',
                      active: props.findReplaceOpen,
                      accent: true,
                      onClick: () => {
                        openPopover.value = 'none';
                        props.onToggleFindReplace();
                      },
                    }),
                  ]),

                  toastMsg.value
                    ? h('span', { class: 'weboffice-qbar__toast', role: 'status' }, toastMsg.value)
                    : null,
                ]
              ),
            ])
          : null,
      ]);
  },
});
