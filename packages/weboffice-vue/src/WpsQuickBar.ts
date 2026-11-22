import {
  Teleport,
  computed,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  type PropType,
} from 'vue';
import {
  useDocxEditor,
  useEditorCommand,
  useEditorState,
  useParagraphFormat,
} from '@docx-editor.dev/vue';
import type {
  Editor,
  EditorSelectionRange,
  EditorSnapshot,
  SupportedImageMime,
} from '@docx-editor.dev/core/contracts/editor';
import {
  DEFAULT_IMAGE_RESOURCE_LIMITS,
  sniffImageMime as sniffImageBytesMime,
  validateRasterHeader,
} from '@docx-editor.dev/core/editor';
import { CN_OFFICIAL_SYMBOLS, charsToTwips } from './units';
import type { OfficialFormatTab } from './OfficialFormatDialog';

type QuickPopover = 'none' | 'lineSpacing' | 'table' | 'symbols' | 'paste';

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

const selectFontSizePt = (snapshot: EditorSnapshot): number => {
  const pt = snapshot.formatting?.fontSizePt;
  return typeof pt === 'number' && pt > 0 ? pt : 16;
};

const selectIsDisabled = (snapshot: EditorSnapshot): boolean =>
  snapshot.isLoading || snapshot.isOpening === true || !snapshot.editable;

interface CutPrecondition {
  readonly editor: Editor;
  readonly loadId: string;
  readonly revision: number;
  readonly range: EditorSelectionRange | null;
  readonly drawingNodeId: string | null;
  readonly selectionKey: string;
  readonly tableKey: string;
  readonly selectedText: string;
}

function sameSelectionRange(
  a: EditorSelectionRange | null,
  b: EditorSelectionRange | null
): boolean {
  if (a === null || b === null) return a === b;
  return (
    a.anchor.blockId === b.anchor.blockId &&
    a.anchor.offset === b.anchor.offset &&
    a.head.blockId === b.head.blockId &&
    a.head.offset === b.head.offset &&
    JSON.stringify(a.scope) === JSON.stringify(b.scope)
  );
}

function captureCutPrecondition(ed: Editor, drawingNodeId: string | null): CutPrecondition | null {
  try {
    const handle = ed.getDocumentHandle();
    const snap = ed.snapshot();
    if (!handle || !handle.loadId || !snap.editable || snap.isLoading || snap.isOpening) {
      return null;
    }
    return {
      editor: ed,
      loadId: handle.loadId,
      revision: handle.revision,
      range: ed.getSelectionRange(),
      drawingNodeId,
      selectionKey: JSON.stringify(snap.selection ?? null),
      tableKey: JSON.stringify(snap.table ?? null),
      selectedText: ed.query({ type: 'selectedText' }),
    };
  } catch {
    return null;
  }
}

function matchesCutPrecondition(before: CutPrecondition, currentEditor: Editor | null): boolean {
  if (!currentEditor || currentEditor !== before.editor) return false;
  try {
    const handle = currentEditor.getDocumentHandle();
    const snap = currentEditor.snapshot();
    if (!handle || !snap.editable || snap.isLoading || snap.isOpening) return false;
    if (handle.loadId !== before.loadId || handle.revision !== before.revision) return false;
    if (!sameSelectionRange(before.range, currentEditor.getSelectionRange())) return false;
    if (JSON.stringify(snap.selection ?? null) !== before.selectionKey) return false;
    if (JSON.stringify(snap.table ?? null) !== before.tableKey) return false;
    const afterFlavours = currentEditor.query({ type: 'copyFlavours' });
    if ((afterFlavours?.drawingNodeId ?? null) !== before.drawingNodeId) return false;
    if (currentEditor.query({ type: 'selectedText' }) !== before.selectedText) return false;
    return currentEditor.can({ type: 'cut' }).ok;
  } catch {
    return false;
  }
}

function isSupportedRasterMime(mime: string | null): mime is SupportedImageMime {
  return mime === 'image/png' || mime === 'image/jpeg' || mime === 'image/gif';
}

function isClipboardPermissionError(err: unknown): boolean {
  const name = (err as { name?: unknown } | null)?.name;
  return name === 'NotAllowedError' || name === 'SecurityError';
}

async function writeClipboardRichChecked(text: string, html: string | null): Promise<boolean> {
  if (!text && !html) return false;
  const clipboard = navigator.clipboard;
  if (!clipboard) return false;
  if (html && typeof clipboard.write === 'function' && typeof ClipboardItem !== 'undefined') {
    try {
      const item = new ClipboardItem({
        'text/plain': new Blob([text], { type: 'text/plain' }),
        'text/html': new Blob([html], { type: 'text/html' }),
      });
      await clipboard.write([item]);
      return true;
    } catch (err) {
      if (isClipboardPermissionError(err)) return false;
      // Only fall back to plain writeText when ClipboardItem/HTML MIME itself is unsupported.
    }
  }
  if (typeof clipboard.writeText === 'function') {
    try {
      await clipboard.writeText(text);
      return true;
    } catch {
      return false;
    }
  }
  return false;
}

async function readClipboardPayload(
  plainOnly: boolean
): Promise<{ text: string; html: string | null }> {
  const clipboard = navigator.clipboard;
  if (!clipboard) throw new Error('clipboard unavailable');
  if (!plainOnly && typeof clipboard.read === 'function' && typeof ClipboardItem !== 'undefined') {
    try {
      const items = await clipboard.read();
      let text = '';
      let html: string | null = null;
      for (const item of items) {
        if (item.types.includes('text/plain')) {
          text = await (await item.getType('text/plain')).text();
        }
        if (item.types.includes('text/html')) {
          html = await (await item.getType('text/html')).text();
        }
      }
      if (text || html) return { text, html };
    } catch (err) {
      if (isClipboardPermissionError(err)) throw err;
      // Fall through to readText when async clipboard.read() is unavailable or unsupported.
    }
  }
  if (typeof clipboard.readText === 'function') {
    return { text: await clipboard.readText(), html: null };
  }
  if (typeof clipboard.read === 'function') {
    const items = await clipboard.read();
    for (const item of items) {
      if (item.types.includes('text/plain')) {
        return { text: await (await item.getType('text/plain')).text(), html: null };
      }
    }
  }
  throw new Error('clipboard read unavailable');
}

/**
 * WPS 风格公文快捷增强工具栏（挂载于主工具栏下方，补齐 WPS 顶部栏高频 P0 能力）：
 * - 剪贴板：复制 / 剪切 / 粘贴 / 无格式粘贴
 * - 段落平铺对齐（左/中/右/两端） + 显示/隐藏段落标记 ¶
 * - 公文快捷排版：固定值行距（28/29/30 磅） / 首行缩进 2 字符一键切换 / 段落与页面设置（厘米/字符）
 * - 快捷插入：可视化表格网格 / 本地图片 / 公文常用符号
 * - 查找与替换直达入口
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
    const imageInsertCmd = useEditorCommand('image.insert');
    // oxlint-disable-next-line react/rules-of-hooks
    const paragraph = useParagraphFormat();
    // oxlint-disable-next-line react/rules-of-hooks
    const currentFontSizePt = useEditorState(selectFontSizePt);
    // oxlint-disable-next-line react/rules-of-hooks
    const isDisabled = useEditorState(selectIsDisabled);

    const anchorRef = ref<HTMLElement | null>(null);
    const chromeBandEl = shallowRef<HTMLElement | null>(null);
    const barRef = ref<HTMLElement | null>(null);
    const fileInputRef = ref<HTMLInputElement | null>(null);
    const openPopover = ref<QuickPopover>('none');
    const tableHover = ref({ rows: 3, cols: 4 });
    const toastMsg = ref('');
    const imageInsertBusy = ref(false);
    let cutSeq = 0;
    let toastTimer: number | null = null;
    let observer: MutationObserver | undefined;

    /**
     * 精确区分「首行缩进 2 字符」与「1 字符 / 其它缩进 / 混选 / 未知」：
     * 仅当选区非混选、已知缩进且与当前字号下 2 字符宽度误差 ≤ 20 twips（1pt）时才高亮为已开启。
     */
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

    onMounted(() => {
      document.addEventListener('pointerdown', onDocPointerDown, true);
      locateChromeBand();
      const host = props.hostElement ?? anchorRef.value?.closest<HTMLElement>('.weboffice-host');
      if (host) {
        observer = new MutationObserver(locateChromeBand);
        observer.observe(host, { childList: true, subtree: true });
      }
    });

    onBeforeUnmount(() => {
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      observer?.disconnect();
      if (toastTimer !== null) window.clearTimeout(toastTimer);
    });

    function togglePopover(next: QuickPopover) {
      openPopover.value = openPopover.value === next ? 'none' : next;
    }

    async function runCopyOrCut(action: 'copy' | 'cut') {
      const ed = editorRef.value;
      if (!ed) return;
      const gate = ed.can({ type: action });
      if (!gate.ok) {
        showToast(action === 'copy' ? '请先选中要复制的内容' : '请先选中要剪切的内容');
        return;
      }
      const flavours = ed.query({ type: 'copyFlavours' });
      if (!flavours || flavours.reason || (!flavours.text && !flavours.html)) {
        showToast(action === 'copy' ? '当前选区不支持复制' : '当前选区不支持剪切');
        return;
      }
      const cutTicket = action === 'cut' ? ++cutSeq : 0;
      const cutPrecondition =
        action === 'cut' ? captureCutPrecondition(ed, flavours.drawingNodeId ?? null) : null;
      if (action === 'cut' && !cutPrecondition) {
        showToast('当前文档状态不稳定，请重新剪切');
        return;
      }
      const wrote = await writeClipboardRichChecked(flavours.text, flavours.html);
      if (!wrote) {
        showToast(
          action === 'copy'
            ? '剪贴板写入被拒绝，请使用 Ctrl+C 复制'
            : '剪贴板写入被拒绝，未删除所选内容（请使用 Ctrl+X）'
        );
        if (editorRef.value === ed) ed.focus();
        return;
      }
      if (action === 'cut') {
        // 严格核对文档身份、修订号、完整选区锚点及并发票据；任一变化均取消删除，绝不误删后选内容。
        if (
          cutTicket !== cutSeq ||
          !cutPrecondition ||
          !matchesCutPrecondition(cutPrecondition, editorRef.value)
        ) {
          showToast('选区或文档已变化，内容已复制但未删除');
          if (editorRef.value === ed) ed.focus();
          return;
        }
        const delRes = ed.exec({ type: 'deleteText' });
        if (!delRes.ok) {
          showToast(`剪切失败：${delRes.reason || '当前位置不可编辑'}`);
          if (editorRef.value === ed) ed.focus();
          return;
        }
      }
      showToast(action === 'copy' ? '已复制' : '已剪切');
      if (editorRef.value === ed) ed.focus();
    }

    async function runPaste(plainOnly: boolean) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      if (!navigator.clipboard) {
        showToast('请直接使用快捷键 Ctrl+V / Ctrl+Shift+V 粘贴');
        ed.focus();
        return;
      }
      try {
        const { text, html } = await readClipboardPayload(plainOnly);
        if (!text && !html) {
          showToast('剪贴板为空');
          ed.focus();
          return;
        }
        const res = plainOnly
          ? ed.exec({ type: 'pasteWithoutFormatting', text })
          : ed.exec({ type: 'paste', text, ...(html ? { html } : {}) });
        if (!res.ok) {
          showToast(`粘贴失败：${res.reason || '当前位置不可编辑'}`);
        }
      } catch {
        showToast('浏览器限制读取剪贴板，请按 Ctrl+V / Ctrl+Shift+V');
      }
      ed.focus();
    }

    function applyLineSpacingPreset(rule: 'exact' | 'multiple', value: number) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'setLineSpacing', rule, value });
      if (!res.ok) {
        showToast('无法设置行距');
      }
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

    function insertTable(rows: number, cols: number) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'insertTable', rows, cols });
      if (!res.ok) {
        showToast('当前位置无法插入表格');
      }
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
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        const limits = DEFAULT_IMAGE_RESOURCE_LIMITS;
        if (bytes.byteLength === 0) {
          showToast('图片文件为空');
          return;
        }
        if (bytes.byteLength > limits.maxEncodedBytes) {
          showToast('图片体积超出限制');
          return;
        }
        const sniffed = sniffImageBytesMime(bytes);
        if (!isSupportedRasterMime(sniffed)) {
          showToast('仅支持 PNG / JPG / GIF 图片');
          return;
        }
        const header = validateRasterHeader(bytes, sniffed);
        if (!header) {
          showToast('图片文件头无效或已损坏');
          return;
        }
        if (
          header.pixelWidth * header.pixelHeight > limits.maxPixels ||
          header.pixelWidth > limits.maxDimension ||
          header.pixelHeight > limits.maxDimension
        ) {
          showToast('图片分辨率超出限制');
          return;
        }
        const dpiX = header.dpiX ?? 96;
        const dpiY = header.dpiY ?? 96;
        const widthPoints = Math.max(1, Math.round(((header.pixelWidth * 72) / dpiX) * 100) / 100);
        const heightPoints = Math.max(
          1,
          Math.round(((header.pixelHeight * 72) / dpiY) * 100) / 100
        );
        const res = await ed.executeImageCommand({
          type: 'insertImage',
          data: bytes,
          mime: sniffed,
          widthPoints,
          heightPoints,
        });
        if (!res.ok) {
          showToast(`插入图片失败：${res.reason || '当前位置不支持'}`);
        } else {
          showToast('已插入图片');
        }
      } catch {
        showToast('读取图片失败');
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
      if (!res.ok) {
        showToast('当前位置无法插入符号');
      }
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
                        ? h('div', { class: 'weboffice-qbar__menu' }, [
                            h(
                              'button',
                              {
                                type: 'button',
                                class: 'weboffice-qbar__menu-item',
                                onMousedown: keepFocus,
                                onClick: () => void runPaste(false),
                              },
                              '粘贴 (Ctrl+V)'
                            ),
                            h(
                              'button',
                              {
                                type: 'button',
                                class: 'weboffice-qbar__menu-item',
                                onMousedown: keepFocus,
                                onClick: () => void runPaste(true),
                              },
                              '仅保留文本 (Ctrl+Shift+V)'
                            ),
                          ])
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

                  // 3. 公文快捷排版组
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
                        ? h('div', { class: 'weboffice-qbar__menu' }, [
                            ...LINE_SPACING_PRESETS.map((preset) =>
                              h(
                                'button',
                                {
                                  key: preset.testId,
                                  type: 'button',
                                  class: 'weboffice-qbar__menu-item',
                                  'data-testid': preset.testId,
                                  onMousedown: keepFocus,
                                  onClick: () => applyLineSpacingPreset(preset.rule, preset.value),
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
                                onClick: () => {
                                  openPopover.value = 'none';
                                  props.onOpenOfficialDialog('paragraph');
                                },
                              },
                              '段落排版设置（厘米 / 字符）…'
                            ),
                          ])
                        : null,
                    ]),
                    qBtn({
                      label: '首行2字',
                      title: '一键切换当前段落首行缩进 2 字符',
                      testId: 'qbar-first-line-2char',
                      active: hasFirstLineIndent.value,
                      disabled: isDisabled.value,
                      onClick: toggleTwoCharFirstLine,
                    }),
                    qBtn({
                      label: '页面(cm)',
                      title: '页面设置（厘米 / GB/T 9704 公文标准页边距）',
                      testId: 'qbar-page-setup-cm',
                      disabled: isDisabled.value,
                      onClick: () => {
                        openPopover.value = 'none';
                        props.onOpenOfficialDialog('page');
                      },
                    }),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 4. 快捷插入组
                  h('div', { class: 'weboffice-qbar__group' }, [
                    h('div', { class: 'weboffice-qbar__dropdown' }, [
                      qBtn({
                        label: '表格 ▾',
                        title: '插入表格',
                        testId: 'qbar-insert-table',
                        active: openPopover.value === 'table',
                        disabled: !tableInsertCmd.isEnabled.value,
                        onClick: () => togglePopover('table'),
                      }),
                      openPopover.value === 'table'
                        ? h('div', { class: 'weboffice-qbar__table-grid' }, [
                            h(
                              'div',
                              { class: 'weboffice-qbar__table-label' },
                              `插入 ${tableHover.value.rows} × ${tableHover.value.cols} 表格`
                            ),
                            h(
                              'div',
                              { class: 'weboffice-qbar__grid-cells' },
                              Array.from({ length: 6 * 8 }, (_, idx) => {
                                const r = Math.floor(idx / 8) + 1;
                                const c = (idx % 8) + 1;
                                const active =
                                  r <= tableHover.value.rows && c <= tableHover.value.cols;
                                return h('button', {
                                  key: `${r}x${c}`,
                                  type: 'button',
                                  class: ['weboffice-qbar__grid-cell', active ? 'is-on' : ''],
                                  'data-testid': `table-cell-${r}x${c}`,
                                  'aria-label': `${r} 行 ${c} 列`,
                                  onMousedown: keepFocus,
                                  onMouseenter: () => {
                                    tableHover.value = { rows: r, cols: c };
                                  },
                                  onClick: () => insertTable(r, c),
                                });
                              })
                            ),
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
                        title: '插入公文常用符号',
                        testId: 'qbar-symbols',
                        active: openPopover.value === 'symbols',
                        disabled: isDisabled.value,
                        onClick: () => togglePopover('symbols'),
                      }),
                      openPopover.value === 'symbols'
                        ? h(
                            'div',
                            { class: 'weboffice-qbar__symbols' },
                            CN_OFFICIAL_SYMBOLS.map((sym) =>
                              h(
                                'button',
                                {
                                  key: sym.char,
                                  type: 'button',
                                  class: 'weboffice-qbar__symbol-btn',
                                  title: sym.label,
                                  onMousedown: keepFocus,
                                  onClick: () => insertSymbol(sym.char),
                                },
                                sym.char
                              )
                            )
                          )
                        : null,
                    ]),
                  ]),
                  h('span', { class: 'weboffice-qbar__sep' }),

                  // 5. 查找替换
                  h('div', { class: 'weboffice-qbar__group' }, [
                    qBtn({
                      label: '查找替换',
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
