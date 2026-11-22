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
} from 'vue';
import { useDocxEditor, useEditorValueCommand } from '@docx-editor.dev/vue';
import { CN_FONT_SIZES, type ChineseFontSizeOption } from './fonts';

let nextListboxId = 1;

/**
 * 政务公文 / WPS 中文字号控件（初号～八号）：
 * 自动定位 EigenPal 工具栏原有字号框并无缝替换为中文字号下拉与步进输入。
 */
export const ChineseFontSize = defineComponent({
  name: 'ChineseFontSize',
  props: {
    sizes: {
      type: Array as () => readonly ChineseFontSizeOption[],
      default: () => CN_FONT_SIZES,
    },
    hostSelector: {
      type: String,
      default: '.weboffice-host, .eigenpal-host',
    },
  },
  setup(props) {
    const anchor = ref<HTMLElement | null>(null);
    const stepper = shallowRef<HTMLElement | null>(null);
    const fontSizeHost = shallowRef<HTMLElement | null>(null);
    const inputRef = ref<HTMLInputElement | null>(null);
    const menuRef = ref<HTMLElement | null>(null);

    // Vue 3 composables inside setup() trigger oxlint's React rules-of-hooks false positive
    // oxlint-disable-next-line react/rules-of-hooks
    const size = useEditorValueCommand('font.size');
    // oxlint-disable-next-line react/rules-of-hooks
    const editor = useDocxEditor();
    const listboxId = `weboffice-cn-fontsize-${nextListboxId++}`;
    const draft = ref<string | null>(null);
    const menuOpen = ref(false);

    const currentHalfPoints = computed(() => size.value.value);

    const displayedSize = computed(() => {
      const value = currentHalfPoints.value;
      if (value === null) return '';
      return (
        props.sizes.find((item) => Math.round(item.pt * 2) === value)?.label ?? String(value / 2)
      );
    });

    watch(currentHalfPoints, () => {
      draft.value = null;
    });

    watch(menuOpen, async (open) => {
      if (!open) return;
      await nextTick();
      menuRef.value
        ?.querySelector<HTMLElement>('[data-selected="true"]')
        ?.scrollIntoView({ block: 'nearest' });
    });

    let observer: MutationObserver | undefined;

    function onOutsidePointerDown(event: PointerEvent) {
      if (!menuOpen.value) return;
      const target = event.target instanceof Node ? event.target : null;
      if (target && (fontSizeHost.value?.contains(target) || menuRef.value?.contains(target)))
        return;
      menuOpen.value = false;
    }

    onMounted(() => {
      document.addEventListener('pointerdown', onOutsidePointerDown, true);
      const rootHost = anchor.value?.closest(props.hostSelector) ?? anchor.value?.parentElement;
      const toolbar = rootHost?.querySelector('.docx-toolbar');
      if (!toolbar) return;

      const locateStepper = () => {
        const nativeInput = toolbar.querySelector('.docx-toolbar__font-size-input');
        stepper.value = nativeInput?.closest<HTMLElement>('.docx-toolbar__stepper') ?? null;
        fontSizeHost.value =
          nativeInput?.closest<HTMLElement>('.docx-toolbar__font-size') ?? stepper.value;
      };
      locateStepper();
      observer = new MutationObserver(locateStepper);
      observer.observe(toolbar, { childList: true, subtree: true });
    });

    onBeforeUnmount(() => {
      document.removeEventListener('pointerdown', onOutsidePointerDown, true);
      observer?.disconnect();
    });

    function commitHalfPoints(halfPoints: number) {
      if (
        Number.isFinite(halfPoints) &&
        halfPoints >= 2 &&
        halfPoints <= 3276 &&
        halfPoints !== size.value.value
      ) {
        size.execute(halfPoints);
      }
    }

    function applySize() {
      const input = draft.value?.trim();
      draft.value = null;
      if (!input || !size.isEnabled.value) return;
      const points =
        props.sizes.find((item) => item.label === input)?.pt ??
        Number(input.replace(/\s*(?:pt|磅)$/i, ''));
      commitHalfPoints(Math.round(points * 2));
    }

    function pickSize(pt: number) {
      draft.value = null;
      menuOpen.value = false;
      if (size.isEnabled.value) {
        commitHalfPoints(Math.round(pt * 2));
      }
      inputRef.value?.blur();
      editor.value?.focus();
    }

    function isSelected(pt: number): boolean {
      return currentHalfPoints.value !== null && Math.round(pt * 2) === currentHalfPoints.value;
    }

    return () =>
      h('span', { ref: anchor, class: 'weboffice-cn-fontsize-anchor', style: 'display:none' }, [
        fontSizeHost.value
          ? h(Teleport, { to: fontSizeHost.value }, [
              h('div', { class: 'weboffice-cn-fontsize' }, [
                h('input', {
                  ref: inputRef,
                  type: 'text',
                  class: 'weboffice-cn-fontsize__input',
                  role: 'combobox',
                  'aria-label': '字号',
                  'aria-expanded': menuOpen.value ? 'true' : 'false',
                  'aria-controls': listboxId,
                  value: draft.value ?? displayedSize.value,
                  disabled: !size.isEnabled.value,
                  onFocus: (e: FocusEvent) => {
                    (e.target as HTMLInputElement).select();
                    menuOpen.value = true;
                  },
                  onBlur: (e: FocusEvent) => {
                    if (
                      e.relatedTarget instanceof Node &&
                      (fontSizeHost.value?.contains(e.relatedTarget) ||
                        menuRef.value?.contains(e.relatedTarget))
                    ) {
                      return;
                    }
                    applySize();
                    menuOpen.value = false;
                  },
                  onInput: (e: Event) => {
                    draft.value = (e.target as HTMLInputElement).value;
                  },
                  onKeydown: (e: KeyboardEvent) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      applySize();
                      menuOpen.value = false;
                      inputRef.value?.blur();
                      editor.value?.focus();
                    } else if (e.key === 'Escape') {
                      e.preventDefault();
                      draft.value = null;
                      menuOpen.value = false;
                      inputRef.value?.blur();
                    }
                  },
                }),
                menuOpen.value
                  ? h(
                      'div',
                      {
                        ref: menuRef,
                        id: listboxId,
                        role: 'listbox',
                        class: 'weboffice-cn-fontsize__menu',
                      },
                      props.sizes.map((item) =>
                        h(
                          'button',
                          {
                            key: item.label,
                            type: 'button',
                            role: 'option',
                            class: 'weboffice-cn-fontsize__option',
                            'data-selected': isSelected(item.pt) ? 'true' : 'false',
                            onMousedown: (e: MouseEvent) => {
                              e.preventDefault();
                              pickSize(item.pt);
                            },
                          },
                          [
                            h('span', { class: 'weboffice-cn-fontsize__label' }, item.label),
                            h('span', { class: 'weboffice-cn-fontsize__pt' }, `${item.pt}pt`),
                          ]
                        )
                      )
                    )
                  : null,
              ]),
            ])
          : null,
      ]);
  },
});
