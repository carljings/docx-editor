import {
  Teleport,
  computed,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
  type PropType,
} from 'vue';
import { useDocxEditor, useEditorEvent, useEditorState, useZoom } from '@docx-editor.dev/vue';
import type { EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { countDocumentWords, type DocumentWordStats } from './units';

const EMPTY_STATS: DocumentWordStats = Object.freeze({
  totalCharsNoSpace: 0,
  totalCharsWithSpace: 0,
  cjkChars: 0,
  latinWords: 0,
  paragraphCount: 0,
});

const ZOOM_PRESETS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2] as const;

const selectPageTotal = (snapshot: EditorSnapshot): number =>
  snapshot.isOpening === true ? 0 : snapshot.page.total;

const selectMode = (snapshot: EditorSnapshot): string =>
  snapshot.editingMode ?? (snapshot.editable ? 'editing' : 'viewing');

/**
 * WPS 风格底部状态栏：
 * - 左侧：当前页 / 总页数、实时字数统计（点击弹出中文/西文/段落详细统计）、编辑模式与 WPS 兼容标识；
 * - 右侧：缩放 `-` / 百分比与适应页宽选择 / `+`。
 */
export const WebOfficeStatusBar = defineComponent({
  name: 'WebOfficeStatusBar',
  props: {
    hostElement: {
      type: Object as PropType<HTMLElement | null>,
      default: null,
    },
    wpsCompatActive: {
      type: Boolean,
      default: false,
    },
  },
  setup(props) {
    // oxlint-disable-next-line react/rules-of-hooks
    const editorRef = useDocxEditor();
    // oxlint-disable-next-line react/rules-of-hooks
    const totalPages = useEditorState(selectPageTotal);
    // oxlint-disable-next-line react/rules-of-hooks
    const editingMode = useEditorState(selectMode);
    // oxlint-disable-next-line react/rules-of-hooks
    const zoom = useZoom();

    const currentPage = ref(1);
    const stats = ref<DocumentWordStats>(EMPTY_STATS);
    const statsOpen = ref(false);
    const barRef = ref<HTMLElement | null>(null);
    const wordBtnRef = ref<HTMLButtonElement | null>(null);
    const popoverRef = ref<HTMLElement | null>(null);
    const popoverPos = ref<{ left: number; bottom: number; width: number }>({
      left: 8,
      bottom: 34,
      width: 210,
    });
    let statsTimer: number | null = null;
    let scrollEl: HTMLElement | null = null;

    function updatePopoverPosition() {
      const host = props.hostElement;
      const btn = wordBtnRef.value;
      if (!host || !btn) return;
      const hostRect = host.getBoundingClientRect();
      const btnRect = btn.getBoundingClientRect();
      const width = Math.min(220, Math.max(168, Math.floor(hostRect.width - 16)));
      const rawLeft = btnRect.left - hostRect.left;
      const maxLeft = Math.max(8, hostRect.width - width - 8);
      const left = Math.max(8, Math.min(rawLeft, maxLeft));
      const bottom = Math.max(32, hostRect.bottom - btnRect.top + 6);
      popoverPos.value = {
        left: Math.round(left),
        bottom: Math.round(bottom),
        width: Math.round(width),
      };
    }

    function recomputeStats() {
      const ed = editorRef.value;
      if (!ed) {
        stats.value = EMPTY_STATS;
        return;
      }
      try {
        stats.value = countDocumentWords(ed.getParagraphs());
      } catch {
        // Editor may still be loading
      }
    }

    function scheduleStats() {
      if (statsTimer !== null) window.clearTimeout(statsTimer);
      statsTimer = window.setTimeout(() => {
        statsTimer = null;
        recomputeStats();
      }, 220);
    }

    function syncCurrentPage() {
      const ed = editorRef.value;
      if (!ed) return;
      try {
        const p = ed.getCurrentPage('viewport');
        if (Number.isFinite(p) && p >= 1) {
          currentPage.value = p;
        }
      } catch {
        // Ignore transient viewport reads during open
      }
    }

    // oxlint-disable-next-line react/rules-of-hooks
    useEditorEvent('change', () => {
      scheduleStats();
      syncCurrentPage();
    });

    function bindViewportScroll() {
      if (scrollEl) {
        scrollEl.removeEventListener('scroll', syncCurrentPage);
        scrollEl = null;
      }
      const host = props.hostElement;
      if (!host) return;
      const vp = host.querySelector<HTMLElement>(
        '.docx-editor__scroll-container, [data-testid="docx-editor-scroll"]'
      );
      if (vp) {
        scrollEl = vp;
        vp.addEventListener('scroll', syncCurrentPage, { passive: true });
      }
    }

    watch(
      [editorRef, totalPages, () => props.hostElement],
      ([ed]) => {
        if (!ed) return;
        recomputeStats();
        bindViewportScroll();
        syncCurrentPage();
      },
      { immediate: true, flush: 'post' }
    );

    function onOutsideClick(event: PointerEvent) {
      if (!statsOpen.value) return;
      const target = event.target instanceof Node ? event.target : null;
      if (target && (barRef.value?.contains(target) || popoverRef.value?.contains(target))) return;
      statsOpen.value = false;
    }

    function onWindowResize() {
      if (statsOpen.value) updatePopoverPosition();
    }

    function onKeydown(event: KeyboardEvent) {
      if (event.key === 'Escape' && statsOpen.value) {
        statsOpen.value = false;
      }
    }

    onMounted(() => {
      document.addEventListener('pointerdown', onOutsideClick, true);
      window.addEventListener('resize', onWindowResize);
      window.addEventListener('keydown', onKeydown);
      bindViewportScroll();
    });

    onBeforeUnmount(() => {
      document.removeEventListener('pointerdown', onOutsideClick, true);
      window.removeEventListener('resize', onWindowResize);
      window.removeEventListener('keydown', onKeydown);
      if (statsTimer !== null) window.clearTimeout(statsTimer);
      if (scrollEl) scrollEl.removeEventListener('scroll', syncCurrentPage);
    });

    const modeLabel = computed(() => {
      if (editingMode.value === 'suggesting') return '修订建议';
      if (editingMode.value === 'viewing') return '只读预览';
      return '编辑模式';
    });

    const zoomPercent = computed(() => Math.round(zoom.zoom.value * 100));
    const isAutoZoom = computed(() => {
      const m = zoom.mode.value;
      return m.type === 'fit' && m.maxZoom !== undefined;
    });
    const isFitWidthZoom = computed(() => {
      const m = zoom.mode.value;
      return m.type === 'fit' && m.maxZoom === undefined;
    });
    const zoomSelectValue = computed(() => {
      if (isAutoZoom.value) return 'auto';
      if (isFitWidthZoom.value) return 'fit';
      return String(zoomPercent.value);
    });

    return () => {
      if (!props.hostElement) return null;
      const total = Math.max(totalPages.value, 1);
      const cur = Math.min(Math.max(currentPage.value, 1), total);

      return h(Teleport, { to: props.hostElement }, [
        statsOpen.value
          ? h(
              'div',
              {
                ref: popoverRef,
                class: 'weboffice-statusbar__popover',
                'data-testid': 'statusbar-wordcount-detail',
                role: 'dialog',
                'aria-label': '文档字数统计',
                style: {
                  left: `${popoverPos.value.left}px`,
                  bottom: `${popoverPos.value.bottom}px`,
                  width: `${popoverPos.value.width}px`,
                },
              },
              [
                h('div', { class: 'weboffice-statusbar__pop-title' }, '文档字数统计'),
                h('div', { class: 'weboffice-statusbar__pop-row' }, [
                  h('span', '总字数（不计空格）'),
                  h('strong', String(stats.value.totalCharsNoSpace)),
                ]),
                h('div', { class: 'weboffice-statusbar__pop-row' }, [
                  h('span', '中文字符与标点'),
                  h('strong', String(stats.value.cjkChars)),
                ]),
                h('div', { class: 'weboffice-statusbar__pop-row' }, [
                  h('span', '英文与数字词数'),
                  h('strong', String(stats.value.latinWords)),
                ]),
                h('div', { class: 'weboffice-statusbar__pop-row' }, [
                  h('span', '字符数（计空格）'),
                  h('strong', String(stats.value.totalCharsWithSpace)),
                ]),
                h('div', { class: 'weboffice-statusbar__pop-row' }, [
                  h('span', '段落数'),
                  h('strong', String(stats.value.paragraphCount)),
                ]),
              ]
            )
          : null,
        h(
          'footer',
          {
            ref: barRef,
            class: 'weboffice-statusbar',
            'data-testid': 'weboffice-statusbar',
            role: 'status',
            onScroll: () => {
              if (statsOpen.value) updatePopoverPosition();
            },
          },
          [
            h('div', { class: 'weboffice-statusbar__left' }, [
              h(
                'span',
                {
                  class: 'weboffice-statusbar__item',
                  'data-testid': 'statusbar-page',
                },
                `页面：${cur} / ${total}`
              ),
              h('span', { class: 'weboffice-statusbar__div' }),
              h('div', { class: 'weboffice-statusbar__stats-wrap' }, [
                h(
                  'button',
                  {
                    ref: wordBtnRef,
                    type: 'button',
                    class: 'weboffice-statusbar__btn',
                    'data-testid': 'statusbar-wordcount',
                    'aria-expanded': statsOpen.value,
                    title: '点击查看详细字数统计',
                    onClick: () => {
                      recomputeStats();
                      const next = !statsOpen.value;
                      if (next) updatePopoverPosition();
                      statsOpen.value = next;
                    },
                  },
                  `字数：${stats.value.totalCharsNoSpace}`
                ),
              ]),
              h('span', { class: 'weboffice-statusbar__div' }),
              h('span', { class: 'weboffice-statusbar__item' }, modeLabel.value),
              props.wpsCompatActive
                ? h(
                    'span',
                    {
                      class: 'weboffice-statusbar__badge',
                      title: '已开启 WPS 中国版公文排版兼容模式',
                    },
                    'WPS 兼容'
                  )
                : null,
            ]),
            h('div', { class: 'weboffice-statusbar__right' }, [
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-statusbar__zoom-btn',
                  'data-testid': 'statusbar-zoom-out',
                  title: '缩小',
                  disabled: !zoom.canZoomOut.value,
                  onClick: () => zoom.zoomOut(),
                },
                '−'
              ),
              h(
                'select',
                {
                  class: 'weboffice-statusbar__zoom-select',
                  'data-testid': 'statusbar-zoom-select',
                  'aria-label': '显示缩放比例',
                  value: zoomSelectValue.value,
                  onChange: (e: Event) => {
                    const val = (e.target as HTMLSelectElement).value;
                    if (val === 'auto') {
                      zoom.auto();
                      return;
                    }
                    if (val === 'fit') {
                      zoom.fitToWidth();
                      return;
                    }
                    const pct = Number(val);
                    if (Number.isFinite(pct) && pct > 0) {
                      zoom.setMode({ type: 'fixed' });
                      zoom.setZoom(pct / 100);
                    }
                  },
                },
                [
                  h(
                    'option',
                    { value: 'auto' },
                    isAutoZoom.value ? `自动 ${zoomPercent.value}%` : '自动适应'
                  ),
                  h(
                    'option',
                    { value: 'fit' },
                    isFitWidthZoom.value ? `适合页宽 ${zoomPercent.value}%` : '适合页宽'
                  ),
                  ...ZOOM_PRESETS.map((z) => {
                    const pct = Math.round(z * 100);
                    return h('option', { key: pct, value: String(pct) }, `${pct}%`);
                  }),
                  !zoom.isFit.value &&
                  !ZOOM_PRESETS.some((z) => Math.round(z * 100) === zoomPercent.value)
                    ? h('option', { value: String(zoomPercent.value) }, `${zoomPercent.value}%`)
                    : null,
                ]
              ),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-statusbar__zoom-btn',
                  'data-testid': 'statusbar-zoom-in',
                  title: '放大',
                  disabled: !zoom.canZoomIn.value,
                  onClick: () => zoom.zoomIn(),
                },
                '+'
              ),
            ]),
          ]
        ),
      ]);
    };
  },
});
