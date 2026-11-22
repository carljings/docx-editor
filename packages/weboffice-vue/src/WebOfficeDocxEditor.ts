import { computed, defineComponent, h, ref, watchEffect, type PropType } from 'vue';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef } from '@docx-editor.dev/vue';
import { zhCN, type Translations } from '@docx-editor.dev/i18n';
import type { DocumentSource } from '@docx-editor.dev/core/contracts/editor';
import { setWpsCompatMode, type WpsCompatOptions } from '@docx-editor.dev/core/layout';
import { installLegacyPolyfills } from './polyfills';
import {
  createOfficialFontResolver,
  type FontDiagnosticsSnapshot,
  type OfficialFontResolverOptions,
} from './fonts';
import { ChineseFontSize } from './ChineseFontSize';
import { createWebOfficeAdapter } from './adapter';

/**
 * 纯 Apache-2.0 的业务级 Vue 3 Word 编辑器封装：
 * - 自动注入国产化老内核浏览器 Polyfill
 * - 默认启用简体中文（zh-CN）与公文中文字号（初号～八号）
 * - 默认启用纯内网离线公文字体解析器（默认禁用无授权的本机 WPS 字体）
 * - 通过 ref 暴露引擎无关的 WebOfficeEditorHandle（load/save/insertTextAtCursor/getDiagnostics）
 */
export const WebOfficeDocxEditor = defineComponent({
  name: 'WebOfficeDocxEditor',
  inheritAttrs: false,
  props: {
    document: {
      type: [Object, String, Uint8Array, ArrayBuffer] as PropType<DocumentSource | undefined>,
      default: undefined,
    },
    mode: {
      type: String as PropType<DocxEditorProps['mode']>,
      default: 'edit',
    },
    title: {
      type: String,
      default: undefined,
    },
    locale: {
      type: String,
      default: 'zh-CN',
    },
    i18n: {
      type: Object as PropType<Translations>,
      default: () => zhCN,
    },
    fonts: {
      type: [Object, Function] as PropType<DocxEditorProps['fonts']>,
      default: undefined,
    },
    fontOptions: {
      type: Object as PropType<OfficialFontResolverOptions>,
      default: undefined,
    },
    chineseFontSize: {
      type: Boolean,
      default: true,
    },
    rulers: {
      type: Boolean,
      default: true,
    },
    chrome: {
      type: Boolean,
      default: true,
    },
    onFontDiagnostics: {
      type: Function as PropType<(snapshot: FontDiagnosticsSnapshot) => void>,
      default: undefined,
    },
    /**
     * 可选 WPS 中国版公文排版兼容模式（如左对齐中文段标点压缩）。
     * 默认 `false` 保持与 MS Word 上游排版一致；需要对齐 WPS 打印基准时可显式开启。
     */
    wpsCompat: {
      type: [Boolean, Object] as PropType<boolean | WpsCompatOptions>,
      default: false,
    },
  },
  setup(props, { attrs, slots, expose }) {
    installLegacyPolyfills();

    watchEffect(() => {
      setWpsCompatMode(props.wpsCompat);
    });

    const hostEl = ref<HTMLElement | null>(null);
    const editorRef = ref<DocxEditorRef | null>(null);

    const managedFontBundle = computed(() => {
      if (props.fonts) return null;
      return createOfficialFontResolver({
        ...props.fontOptions,
        onDiagnostics: (snap) => {
          props.fontOptions?.onDiagnostics?.(snap);
          props.onFontDiagnostics?.(snap);
        },
      });
    });

    const resolvedFonts = computed(() => props.fonts ?? managedFontBundle.value?.resolver);

    const adapter = createWebOfficeAdapter({
      getRef: () => editorRef.value,
      getHostElement: () => hostEl.value,
      getFontDiagnostics: () => managedFontBundle.value?.getDiagnostics(),
    });

    expose(adapter);

    return () => {
      const { class: attrClass, ...restAttrs } = attrs;
      return h(
        'div',
        {
          ref: hostEl,
          class: [
            'weboffice-host',
            props.chineseFontSize ? 'weboffice-host--cn-fontsize' : '',
            attrClass,
          ],
        },
        [
          h(
            DocxEditor,
            {
              ...restAttrs,
              ref: editorRef,
              document: props.document,
              mode: props.mode,
              title: props.title,
              locale: props.locale,
              i18n: props.i18n,
              fonts: resolvedFonts.value,
              rulers: props.rulers,
              chrome: props.chrome,
            },
            {
              ...slots,
              default: () => [
                props.chineseFontSize ? h(ChineseFontSize) : null,
                slots.default ? slots.default() : null,
              ],
            }
          ),
        ]
      );
    };
  },
});
