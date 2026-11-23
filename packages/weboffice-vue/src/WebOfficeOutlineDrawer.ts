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
import {
  useDocxEditor,
  useDocumentOutline,
  useEditorEvent,
  useEditorState,
} from '@docx-editor.dev/vue';
import type { BookmarkSummary, EditorSnapshot } from '@docx-editor.dev/core/contracts/editor';
import { extractOfficialFallbackHeadings, type FallbackOfficialHeading } from './p2-presets';

type DrawerTab = 'headings' | 'bookmarks';

const selectReadOnly = (snapshot: EditorSnapshot): boolean =>
  snapshot.isLoading || snapshot.isOpening === true || !snapshot.editable;

/**
 * 左侧可折叠大纲、自动目录（TOC）与书签导航抽屉：
 * - 大纲页签：优先显示原生 Word 标题层级（Heading / outlineLvl），缺失时自动识别公文一/二级小节；底部提供插入/更新/移除自动目录入口；
 * - 书签页签：列出全部书签并支持一键跳转、新建与删除（只读模式自动禁用写操作）。
 */
export const WebOfficeOutlineDrawer = defineComponent({
  name: 'WebOfficeOutlineDrawer',
  props: {
    open: {
      type: Boolean,
      required: true,
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
    const outline = useDocumentOutline();
    // oxlint-disable-next-line react/rules-of-hooks
    const isReadOnly = useEditorState(selectReadOnly);

    const activeTab = ref<DrawerTab>('headings');
    const fallbackHeadings = ref<readonly FallbackOfficialHeading[]>([]);
    const bookmarks = ref<readonly BookmarkSummary[]>([]);
    const newBookmarkName = ref('');
    const feedback = ref('');
    const revisionTick = ref(0);
    const drawerLeftPx = ref(0);
    const drawerRef = ref<HTMLElement | null>(null);

    function syncDrawerLeft() {
      const left = props.hostElement?.getBoundingClientRect().left ?? 0;
      drawerLeftPx.value = left < 0 ? Math.ceil(-left) : 0;
    }

    const hasNativeOutline = computed(() => !outline.isEmpty.value);

    function refreshDrawerData() {
      revisionTick.value += 1;
      const ed = editorRef.value;
      if (!ed) {
        fallbackHeadings.value = [];
        bookmarks.value = [];
        return;
      }
      try {
        bookmarks.value = ed.query({ type: 'bookmarks' }) ?? [];
      } catch {
        bookmarks.value = [];
      }
      if (outline.isEmpty.value) {
        try {
          fallbackHeadings.value = extractOfficialFallbackHeadings(ed.getParagraphs());
        } catch {
          fallbackHeadings.value = [];
        }
      } else {
        fallbackHeadings.value = [];
      }
    }

    let lastLoadId: string | null = null;
    function syncLoadIdentity() {
      const nextLoadId = editorRef.value?.getDocumentHandle()?.loadId ?? null;
      if (nextLoadId !== lastLoadId) {
        lastLoadId = nextLoadId;
        activeTab.value = 'headings';
        feedback.value = '';
      }
    }

    // oxlint-disable-next-line react/rules-of-hooks
    useEditorEvent('change', (ev) => {
      if (ev.source === 'load') {
        lastLoadId = editorRef.value?.getDocumentHandle()?.loadId ?? null;
        activeTab.value = 'headings';
        feedback.value = '';
      }
      if (props.open) refreshDrawerData();
    });

    // oxlint-disable-next-line react/rules-of-hooks
    useEditorEvent('selectionChange', () => {
      if (props.open) revisionTick.value += 1;
    });

    watch(
      [() => props.open, () => editorRef.value, () => outline.items.value.length],
      ([isOpen]) => {
        syncLoadIdentity();
        if (isOpen) {
          feedback.value = '';
          syncDrawerLeft();
          refreshDrawerData();
        }
      },
      { immediate: true }
    );

    function onKeydown(event: KeyboardEvent) {
      if (!props.open || event.key !== 'Escape' || event.isComposing) return;
      event.preventDefault();
      props.onClose();
      editorRef.value?.focus();
    }

    onMounted(() => {
      document.addEventListener('keydown', onKeydown);
    });

    onBeforeUnmount(() => {
      document.removeEventListener('keydown', onKeydown);
    });

    function jumpFallbackHeading(blockId: string) {
      const ed = editorRef.value;
      if (!ed) return;
      ed.scrollToBlock(blockId, { block: 'center', behavior: 'instant' });
      ed.focus();
    }

    function runTocAction(action: 'insert' | 'refresh' | 'remove') {
      const ed = editorRef.value;
      if (!ed) return;
      const res =
        action === 'insert'
          ? ed.exec({ type: 'insertToc' })
          : action === 'refresh'
            ? ed.exec({ type: 'refreshToc', mode: 'entire' })
            : ed.exec({ type: 'removeToc' });
      if (!res.ok) {
        feedback.value = res.reason || '当前位置无法执行目录操作';
      } else {
        feedback.value =
          action === 'insert'
            ? '已插入自动目录'
            : action === 'refresh'
              ? '已更新目录'
              : '已移除目录';
      }
      ed.focus();
    }

    function addBookmark() {
      const ed = editorRef.value;
      if (!ed) return;
      const name = newBookmarkName.value.trim();
      if (!name) {
        feedback.value = '请输入书签名称（须以字母或汉字开头，不含空格）';
        return;
      }
      const res = ed.exec({ type: 'insertBookmark', name });
      if (!res.ok) {
        feedback.value = res.reason || '书签名称不合法或当前位置不可添加';
        return;
      }
      newBookmarkName.value = '';
      feedback.value = `已添加书签「${name}」`;
      refreshDrawerData();
      ed.focus();
    }

    function goToBookmark(name: string) {
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'goToBookmark', name });
      if (!res.ok) {
        feedback.value = res.reason || '无法定位该书签';
        return;
      }
      feedback.value = `已跳转至书签「${name}」`;
      ed.focus();
    }

    function deleteBookmark(name: string) {
      const ed = editorRef.value;
      if (!ed) return;
      const res = ed.exec({ type: 'deleteBookmark', name });
      if (!res.ok) {
        feedback.value = res.reason || '无法删除该书签';
        return;
      }
      feedback.value = `已删除书签「${name}」`;
      refreshDrawerData();
      ed.focus();
    }

    const canInsertToc = computed(() => {
      void revisionTick.value;
      return !isReadOnly.value && Boolean(editorRef.value?.can({ type: 'insertToc' }).ok);
    });
    const canRefreshToc = computed(() => {
      void revisionTick.value;
      return !isReadOnly.value && Boolean(editorRef.value?.can({ type: 'refreshToc' }).ok);
    });
    const canRemoveToc = computed(() => {
      void revisionTick.value;
      return !isReadOnly.value && Boolean(editorRef.value?.can({ type: 'removeToc' }).ok);
    });

    return () => {
      if (!props.open || !props.hostElement) return null;
      return h(Teleport, { to: props.hostElement }, [
        h(
          'aside',
          {
            ref: drawerRef,
            class: 'weboffice-outline-drawer',
            'data-testid': 'weboffice-outline-drawer',
            'aria-label': '文档大纲与书签导航',
            style: drawerLeftPx.value > 0 ? { left: `${drawerLeftPx.value}px` } : undefined,
          },
          [
            h('header', { class: 'weboffice-outline__head' }, [
              h('div', { class: 'weboffice-outline__tabs', role: 'tablist' }, [
                h(
                  'button',
                  {
                    type: 'button',
                    role: 'tab',
                    'aria-selected': activeTab.value === 'headings',
                    class: [
                      'weboffice-outline__tab',
                      activeTab.value === 'headings' ? 'is-active' : '',
                    ],
                    'data-testid': 'outline-tab-headings',
                    onClick: () => {
                      activeTab.value = 'headings';
                      feedback.value = '';
                      refreshDrawerData();
                    },
                  },
                  '大纲目录'
                ),
                h(
                  'button',
                  {
                    type: 'button',
                    role: 'tab',
                    'aria-selected': activeTab.value === 'bookmarks',
                    class: [
                      'weboffice-outline__tab',
                      activeTab.value === 'bookmarks' ? 'is-active' : '',
                    ],
                    'data-testid': 'outline-tab-bookmarks',
                    onClick: () => {
                      activeTab.value = 'bookmarks';
                      feedback.value = '';
                      refreshDrawerData();
                    },
                  },
                  `书签 (${bookmarks.value.length})`
                ),
              ]),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-outline__close',
                  'data-testid': 'outline-close',
                  'aria-label': '关闭导航面板',
                  onClick: () => {
                    props.onClose();
                    editorRef.value?.focus();
                  },
                },
                '×'
              ),
            ]),

            activeTab.value === 'headings'
              ? h('div', { class: 'weboffice-outline__body' }, [
                  hasNativeOutline.value
                    ? h(
                        'div',
                        { class: 'weboffice-outline__list', 'data-testid': 'outline-heading-list' },
                        outline.items.value.map((item) =>
                          h(
                            'button',
                            {
                              key: item.heading.blockId,
                              type: 'button',
                              class: [
                                'weboffice-outline__item',
                                outline.selectedBlockId.value === item.heading.blockId
                                  ? 'is-active'
                                  : '',
                              ],
                              style: { paddingLeft: `${10 + Math.min(item.depth, 4) * 12}px` },
                              onClick: () => outline.goTo(item.heading.blockId),
                            },
                            item.heading.text || '（无标题）'
                          )
                        )
                      )
                    : fallbackHeadings.value.length > 0
                      ? h(
                          'div',
                          {
                            class: 'weboffice-outline__list',
                            'data-testid': 'outline-fallback-list',
                          },
                          [
                            h('div', { class: 'weboffice-outline__hint' }, '自动识别公文小节'),
                            ...fallbackHeadings.value.map((item) =>
                              h(
                                'button',
                                {
                                  key: item.blockId,
                                  type: 'button',
                                  class: 'weboffice-outline__item',
                                  style: { paddingLeft: `${10 + item.depth * 14}px` },
                                  onClick: () => jumpFallbackHeading(item.blockId),
                                },
                                item.text
                              )
                            ),
                          ]
                        )
                      : h(
                          'div',
                          { class: 'weboffice-outline__empty' },
                          '当前文档尚无标题样式或公文编号小节。应用「标题 1–3」后即可自动生成目录。'
                        ),
                  h('div', { class: 'weboffice-outline__toc-actions' }, [
                    h(
                      'button',
                      {
                        type: 'button',
                        class: 'weboffice-qbar__chip-btn is-accent',
                        'data-testid': 'outline-insert-toc',
                        disabled: !canInsertToc.value,
                        onClick: () => runTocAction('insert'),
                      },
                      '插入目录'
                    ),
                    h(
                      'button',
                      {
                        type: 'button',
                        class: 'weboffice-qbar__chip-btn',
                        'data-testid': 'outline-refresh-toc',
                        disabled: !canRefreshToc.value,
                        onClick: () => runTocAction('refresh'),
                      },
                      '更新目录'
                    ),
                    h(
                      'button',
                      {
                        type: 'button',
                        class: 'weboffice-qbar__chip-btn is-danger',
                        'data-testid': 'outline-remove-toc',
                        disabled: !canRemoveToc.value,
                        onClick: () => runTocAction('remove'),
                      },
                      '移除目录'
                    ),
                  ]),
                ])
              : h('div', { class: 'weboffice-outline__body' }, [
                  h('div', { class: 'weboffice-outline__bm-add' }, [
                    h('input', {
                      type: 'text',
                      class: 'weboffice-qbar__custom-input weboffice-outline__bm-input',
                      'data-testid': 'outline-bookmark-input',
                      placeholder: '新书签名（字母/汉字开头）',
                      disabled: isReadOnly.value,
                      value: newBookmarkName.value,
                      onInput: (e: Event) => {
                        newBookmarkName.value = (e.target as HTMLInputElement).value;
                      },
                      onKeydown: (e: KeyboardEvent) => {
                        if (e.key === 'Enter' && !e.isComposing) {
                          e.preventDefault();
                          addBookmark();
                        }
                      },
                    }),
                    h(
                      'button',
                      {
                        type: 'button',
                        class: 'weboffice-qbar__custom-apply',
                        'data-testid': 'outline-bookmark-add',
                        disabled: isReadOnly.value,
                        onClick: addBookmark,
                      },
                      '添加'
                    ),
                  ]),
                  bookmarks.value.length > 0
                    ? h(
                        'div',
                        {
                          class: 'weboffice-outline__list',
                          'data-testid': 'outline-bookmark-list',
                        },
                        bookmarks.value.map((bm) =>
                          h('div', { key: bm.name, class: 'weboffice-outline__bm-row' }, [
                            h(
                              'button',
                              {
                                type: 'button',
                                class: 'weboffice-outline__item',
                                'data-testid': `bookmark-goto-${bm.name}`,
                                onClick: () => goToBookmark(bm.name),
                              },
                              bm.name
                            ),
                            h(
                              'button',
                              {
                                type: 'button',
                                class: 'weboffice-outline__bm-del',
                                'data-testid': `bookmark-del-${bm.name}`,
                                'aria-label': `删除书签 ${bm.name}`,
                                disabled: isReadOnly.value,
                                onClick: () => deleteBookmark(bm.name),
                              },
                              '×'
                            ),
                          ])
                        )
                      )
                    : h(
                        'div',
                        { class: 'weboffice-outline__empty' },
                        '暂无书签，可在光标处添加书签以便快速定位。'
                      ),
                ]),

            feedback.value
              ? h('div', { class: 'weboffice-outline__feedback', role: 'status' }, feedback.value)
              : null,
          ]
        ),
      ]);
    };
  },
});
