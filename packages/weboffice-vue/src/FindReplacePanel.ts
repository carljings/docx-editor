import {
  Teleport,
  computed,
  defineComponent,
  h,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
  type PropType,
} from 'vue';
import { useDocxEditor, useDocumentSearch } from '@docx-editor.dev/vue';

/**
 * WPS 风格悬浮「查找与替换」面板：
 * - 复用 @docx-editor.dev/vue 的 useDocumentSearch() 共享搜索会话与高亮层；
 * - 补齐单个替换（replaceMatch）与全部替换（replaceAllMatches，单步可撤销）；
 * - 支持在宿主容器内通过 Ctrl+H / Cmd+Shift+H / Cmd+Option+F 呼出，Escape 关闭。
 */
export const FindReplacePanel = defineComponent({
  name: 'FindReplacePanel',
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
    onOpenRequest: {
      type: Function as PropType<() => void>,
      default: undefined,
    },
  },
  setup(props) {
    // oxlint-disable-next-line react/rules-of-hooks
    const editorRef = useDocxEditor();
    // oxlint-disable-next-line react/rules-of-hooks
    const search = useDocumentSearch(() => ({
      highlight: props.open ? 'all' : 'none',
    }));

    const findInputRef = ref<HTMLInputElement | null>(null);
    const replaceInputRef = ref<HTMLInputElement | null>(null);
    const replaceText = ref('');
    const notice = ref('');
    const noticeKind = ref<'info' | 'error'>('info');

    const totalMatches = computed(() => search.matches.value.length);
    const activeIndex = computed(() => search.activeIndex.value);
    const hasQuery = computed(() => search.query.value.length > 0);
    const countLabel = computed(() => {
      if (!hasQuery.value) return '';
      if (search.isPending.value) return '查找中…';
      if (totalMatches.value === 0) return '0 / 0';
      const suffix = search.truncated.value ? '+' : '';
      const current = activeIndex.value >= 0 ? activeIndex.value + 1 : 0;
      return `${current} / ${totalMatches.value}${suffix}`;
    });

    const canReplaceOne = computed(() => {
      const ed = editorRef.value;
      if (!ed || !hasQuery.value || totalMatches.value === 0) return false;
      const target = search.activeMatch.value ?? search.matches.value[0];
      if (!target) return false;
      return ed.can({ type: 'replaceMatch', match: target, text: replaceText.value }).ok;
    });

    const canReplaceAll = computed(() => {
      const ed = editorRef.value;
      if (!ed || !hasQuery.value || totalMatches.value === 0) return false;
      return ed.can({
        type: 'replaceAllMatches',
        query: search.query.value,
        text: replaceText.value,
        matchCase: search.matchCase.value,
        wholeWord: search.wholeWord.value,
      }).ok;
    });

    function showNotice(message: string, kind: 'info' | 'error' = 'info') {
      notice.value = message;
      noticeKind.value = kind;
    }

    // 当面板打开且搜索有结果但尚未定位时，自动定位到第 1 处（不抢占输入框焦点）
    watch(
      [() => props.open, totalMatches, activeIndex, () => search.isPending.value],
      ([isOpen, total, idx, pending]) => {
        if (isOpen && !pending && total > 0 && idx < 0) {
          search.goTo(0);
        }
      }
    );

    watch(
      () => props.open,
      async (isOpen) => {
        notice.value = '';
        if (!isOpen) {
          return;
        }
        // 若用户当前选中了一段单行文字，打开面板时自动填入查找框
        const ed = editorRef.value;
        if (ed) {
          const selected = ed.query({ type: 'selectedText' });
          if (typeof selected === 'string') {
            const trimmed = selected.trim();
            if (trimmed.length > 0 && trimmed.length <= 80 && !trimmed.includes('\n')) {
              search.setQuery(trimmed);
            }
          }
        }
        await nextTick();
        findInputRef.value?.focus();
        findInputRef.value?.select();
      }
    );

    function handleFindStep(direction: 'next' | 'previous') {
      if (!hasQuery.value) return;
      notice.value = '';
      if (search.isPending.value) {
        search.find(search.query.value, { selectFirst: true });
        return;
      }
      if (direction === 'next') search.next();
      else search.previous();
    }

    function handleReplaceCurrent() {
      const ed = editorRef.value;
      if (!ed || !hasQuery.value) return;
      if (search.isPending.value) {
        search.find(search.query.value, { selectFirst: true });
      }
      let target = search.activeMatch.value;
      if (!target) {
        if (search.matches.value.length === 0) {
          showNotice('未找到可替换的匹配项', 'error');
          return;
        }
        search.goTo(0);
        target = search.activeMatch.value ?? search.matches.value[0] ?? null;
      }
      if (!target) return;
      const prevIdx = Math.max(0, search.activeIndex.value);
      const res = ed.exec({ type: 'replaceMatch', match: target, text: replaceText.value });
      if (!res.ok) {
        showNotice(`替换失败：${res.reason || '当前位置不可编辑'}`, 'error');
        return;
      }
      const nextMatches = search.find(search.query.value, { selectFirst: false });
      if (nextMatches.length > 0) {
        const nextIdx = Math.min(prevIdx, nextMatches.length - 1);
        search.goTo(nextIdx);
        showNotice(`已替换 1 处，剩余 ${nextMatches.length} 处`);
      } else {
        showNotice('已替换最后 1 处匹配项');
      }
    }

    function handleReplaceAll() {
      const ed = editorRef.value;
      if (!ed || !hasQuery.value) return;
      if (search.isPending.value) {
        search.find(search.query.value, { selectFirst: false });
      }
      const res = ed.exec({
        type: 'replaceAllMatches',
        query: search.query.value,
        text: replaceText.value,
        matchCase: search.matchCase.value,
        wholeWord: search.wholeWord.value,
      });
      if (!res.ok) {
        const partial = res.replacements;
        if (partial && partial.replaced > 0) {
          showNotice(`已替换 ${partial.replaced} 处，跳过 ${partial.skipped} 处`, 'error');
        } else {
          showNotice(`全部替换未完成：${res.reason || '文档只读或目标受保护'}`, 'error');
        }
        search.find(search.query.value, { selectFirst: false });
        return;
      }
      const outcome = res.replacements;
      const replaced = outcome?.replaced ?? 0;
      const skipped = outcome?.skipped ?? 0;
      search.find(search.query.value, { selectFirst: false });
      if (skipped > 0) {
        showNotice(`已替换 ${replaced} 处，跳过 ${skipped} 处受保护内容`);
      } else if (replaced > 0) {
        showNotice(`已完成全部替换，共替换 ${replaced} 处（可按 Ctrl+Z 撤销）`);
      } else {
        showNotice('未找到可替换的内容');
      }
    }

    function closePanel() {
      search.clear();
      notice.value = '';
      props.onClose();
      editorRef.value?.focus();
    }

    function onGlobalKeydown(event: KeyboardEvent) {
      const host = props.hostElement;
      if (!host) return;
      const activeEl = document.activeElement;
      const insideHost = Boolean(activeEl && host.contains(activeEl));
      const isCtrlOrCmd = event.ctrlKey || event.metaKey;
      const key = event.key.toLowerCase();

      // Ctrl+H（Win/Linux）或 Cmd+Shift+H / Cmd+Option+F（macOS）呼出查找替换
      const openChord =
        (isCtrlOrCmd && key === 'h' && !event.altKey) ||
        (event.metaKey && event.altKey && key === 'f');
      if (openChord && (insideHost || props.open)) {
        event.preventDefault();
        event.stopPropagation();
        props.onOpenRequest?.();
        return;
      }

      if (event.key === 'Escape' && props.open && insideHost) {
        event.preventDefault();
        closePanel();
      }
    }

    onMounted(() => {
      window.addEventListener('keydown', onGlobalKeydown, true);
    });

    onBeforeUnmount(() => {
      window.removeEventListener('keydown', onGlobalKeydown, true);
    });

    return () => {
      if (!props.open || !props.hostElement) return null;
      return h(Teleport, { to: props.hostElement }, [
        h(
          'section',
          {
            class: 'weboffice-find-replace',
            'data-testid': 'weboffice-find-replace',
            role: 'dialog',
            'aria-label': '查找与替换',
          },
          [
            h('header', { class: 'weboffice-find-replace__head' }, [
              h('strong', { class: 'weboffice-find-replace__title' }, '查找与替换'),
              h(
                'span',
                { class: 'weboffice-find-replace__count', 'data-testid': 'fr-count' },
                countLabel.value
              ),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-find-replace__close',
                  title: '关闭 (Esc)',
                  'aria-label': '关闭查找与替换',
                  onClick: closePanel,
                },
                '×'
              ),
            ]),
            h('div', { class: 'weboffice-find-replace__row' }, [
              h('input', {
                ref: findInputRef,
                type: 'text',
                class: 'weboffice-find-replace__input',
                'data-testid': 'fr-find-input',
                placeholder: '查找内容…',
                'aria-label': '查找内容',
                value: search.query.value,
                onInput: (e: Event) => {
                  notice.value = '';
                  search.setQuery((e.target as HTMLInputElement).value);
                },
                onKeydown: (e: KeyboardEvent) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleFindStep(e.shiftKey ? 'previous' : 'next');
                  } else if (e.key === 'Escape') {
                    e.preventDefault();
                    closePanel();
                  }
                },
              }),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-find-replace__icon-btn',
                  'data-testid': 'fr-prev',
                  title: '上一个 (Shift+Enter)',
                  disabled: totalMatches.value === 0,
                  onClick: () => handleFindStep('previous'),
                },
                '↑'
              ),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-find-replace__icon-btn',
                  'data-testid': 'fr-next',
                  title: '下一个 (Enter)',
                  disabled: totalMatches.value === 0,
                  onClick: () => handleFindStep('next'),
                },
                '↓'
              ),
            ]),
            h('div', { class: 'weboffice-find-replace__row' }, [
              h('input', {
                ref: replaceInputRef,
                type: 'text',
                class: 'weboffice-find-replace__input',
                'data-testid': 'fr-replace-input',
                placeholder: '替换为…',
                'aria-label': '替换为',
                value: replaceText.value,
                onInput: (e: Event) => {
                  replaceText.value = (e.target as HTMLInputElement).value;
                },
                onKeydown: (e: KeyboardEvent) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    handleReplaceCurrent();
                  } else if (e.key === 'Escape') {
                    e.preventDefault();
                    closePanel();
                  }
                },
              }),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-find-replace__btn',
                  'data-testid': 'fr-replace-one',
                  disabled: !canReplaceOne.value,
                  onClick: handleReplaceCurrent,
                },
                '替换'
              ),
              h(
                'button',
                {
                  type: 'button',
                  class: 'weboffice-find-replace__btn weboffice-find-replace__btn--primary',
                  'data-testid': 'fr-replace-all',
                  disabled: !canReplaceAll.value,
                  onClick: handleReplaceAll,
                },
                '全部替换'
              ),
            ]),
            h('div', { class: 'weboffice-find-replace__options' }, [
              h('label', { class: 'weboffice-find-replace__check' }, [
                h('input', {
                  type: 'checkbox',
                  checked: search.matchCase.value,
                  onChange: (e: Event) =>
                    search.setMatchCase((e.target as HTMLInputElement).checked),
                }),
                '区分大小写',
              ]),
              h('label', { class: 'weboffice-find-replace__check' }, [
                h('input', {
                  type: 'checkbox',
                  checked: search.wholeWord.value,
                  onChange: (e: Event) =>
                    search.setWholeWord((e.target as HTMLInputElement).checked),
                }),
                '全词匹配',
              ]),
            ]),
            notice.value
              ? h(
                  'div',
                  {
                    class: [
                      'weboffice-find-replace__notice',
                      noticeKind.value === 'error' ? 'weboffice-find-replace__notice--error' : '',
                    ],
                    'data-testid': 'fr-notice',
                    role: 'status',
                  },
                  notice.value
                )
              : null,
          ]
        ),
      ]);
    };
  },
});
