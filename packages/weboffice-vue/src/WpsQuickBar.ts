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
import type { EditorSnapshot, SupportedImageMime } from '@docx-editor.dev/core/contracts/editor';
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

function sniffImageMime(file: File): SupportedImageMime | null {
  if (file.type === 'image/png' || file.type === 'image/jpeg' || file.type === 'image/gif') {
    return file.type;
  }
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.gif')) return 'image/gif';
  return null;
}

function readImageDimensions(blob: Blob): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth || 320;
      const h = img.naturalHeight || 200;
      URL.revokeObjectURL(url);
      resolve({ width: w, height: h });
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve({ width: 320, height: 200 });
    };
    img.src = url;
  });
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
    let toastTimer: number | null = null;
    let observer: MutationObserver | undefined;

    const hasFirstLineIndent = computed(
      () => (paragraph.format.value?.indentFirstLineTwips ?? 0) > 40
    );

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

    function runCopyOrCut(action: 'copy' | 'cut') {
      const ed = editorRef.value;
      if (!ed) return;
      const selected = ed.query({ type: 'selectedText' });
      const res = ed.exec({ type: action });
      if (!res.ok) {
        showToast(action === 'copy' ? '请先选中要复制的内容' : '请先选中要剪切的内容');
        return;
      }
      if (typeof selected === 'string' && selected.length > 0 && navigator.clipboard?.writeText) {
        void navigator.clipboard.writeText(selected).catch(() => {});
      }
      showToast(action === 'copy' ? '已复制' : '已剪切');
      ed.focus();
    }

    async function runPaste(plainOnly: boolean) {
      openPopover.value = 'none';
      const ed = editorRef.value;
      if (!ed) return;
      if (!navigator.clipboard?.readText) {
        showToast('请直接使用快捷键 Ctrl+V / Ctrl+Shift+V 粘贴');
        ed.focus();
        return;
      }
      try {
        const text = await navigator.clipboard.readText();
        if (!text) {
          showToast('剪贴板为空');
          ed.focus();
          return;
        }
        const cmd = plainOnly
          ? ({ type: 'pasteWithoutFormatting', text } as const)
          : ({ type: 'paste', text } as const);
        const res = ed.exec(cmd);
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
      const targetTwips = hasFirstLineIndent.value ? 0 : charsToTwips(2, currentFontSizePt.value);
      const res = ed.exec({ type: 'setIndent', firstLine: targetTwips });
      if (!res.ok) {
        showToast('当前位置无法调整首行缩进');
      } else {
        showToast(targetTwips > 0 ? '已设首行缩进 2 字符' : '已取消首行缩进');
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
      const mime = sniffImageMime(file);
      if (!mime) {
        showToast('仅支持 PNG / JPG / GIF 图片');
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        const dims = await readImageDimensions(file);
        const rawW = (dims.width * 72) / 96;
        const rawH = (dims.height * 72) / 96;
        const maxWidthPt = 420;
        const scale = rawW > maxWidthPt ? maxWidthPt / rawW : 1;
        const res = ed.exec({
          type: 'insertImage',
          data: bytes,
          mime,
          widthPoints: Math.max(24, Math.round(rawW * scale)),
          heightPoints: Math.max(24, Math.round(rawH * scale)),
        });
        if (!res.ok) {
          showToast(`插入图片失败：${res.reason || '当前位置不支持'}`);
        }
      } catch {
        showToast('读取图片失败');
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
                      title: '复制选中文本 (Ctrl+C)',
                      testId: 'qbar-copy',
                      onClick: () => runCopyOrCut('copy'),
                    }),
                    qBtn({
                      label: '剪切',
                      title: '剪切选中文本 (Ctrl+X)',
                      testId: 'qbar-cut',
                      disabled: isDisabled.value,
                      onClick: () => runCopyOrCut('cut'),
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
