import type { DocxEditorRef } from '@docx-editor.dev/vue';
import type { DocumentSource, Editor } from '@docx-editor.dev/core/contracts/editor';
import { probeBrowserEnvironment, type BrowserEnvironmentProbe } from './polyfills';
import type { FontDiagnosticsSnapshot } from './fonts';

export interface CursorPosition {
  readonly blockId: string;
  readonly offset: number;
}

export type InsertTextMode = 'cursor' | 'drop-position' | 'fallback-first-paragraph' | 'failed';

export interface InsertTextOptions {
  readonly dropPosition?: CursorPosition;
  /**
   * 失焦或当前无有效选区时，是否允许兜底写入首段末尾。
   * 业务/T7 严格验收默认 `false`，避免把兜底插入误判为指定光标插入。
   */
  readonly allowFallback?: boolean;
}

export interface InsertTextResult {
  readonly ok: boolean;
  readonly changed: boolean;
  readonly mode: InsertTextMode;
  readonly fallbackApplied: boolean;
  readonly position?: CursorPosition;
  readonly reason?: string;
}

export interface WebOfficeDiagnostics {
  readonly engine: 'eigenpal-apache-core';
  readonly pageCount: number;
  readonly browser: BrowserEnvironmentProbe;
  readonly fonts?: FontDiagnosticsSnapshot;
}

/**
 * 引擎无关的业务适配契约：
 * 报告中心（及未来其他业务系统）优先面向此接口开发，
 * 以便后续切换底层排版内核时宿主侧鉴权、指标插入、保存回调与附件版本逻辑零改造。
 */
export interface WebOfficeEditorHandle {
  load(document: DocumentSource): void;
  save(): Promise<Uint8Array | null>;
  focus(): void;
  insertTextAtCursor(
    text: string,
    dropPositionOrOptions?: CursorPosition | InsertTextOptions
  ): InsertTextResult;
  getPageCount(): number;
  getDiagnostics(): WebOfficeDiagnostics;
  getRawEditor(): Editor | null;
  getRawRef(): DocxEditorRef | null;
}

export function createWebOfficeAdapter(options: {
  getRef: () => DocxEditorRef | null;
  getHostElement?: () => HTMLElement | null;
  getFontDiagnostics?: () => FontDiagnosticsSnapshot | undefined;
}): WebOfficeEditorHandle {
  const getPageCount = (): number => {
    const host = options.getHostElement?.();
    if (host) {
      const domPages = host.querySelectorAll('.docx-page').length;
      if (domPages > 0) return domPages;
    }
    const ed = options.getRef()?.getEditor() as (Editor & { getPageCount?: () => number }) | null;
    if (ed && typeof ed.getPageCount === 'function') {
      try {
        return ed.getPageCount();
      } catch {
        return 0;
      }
    }
    return 0;
  };

  return {
    load(document: DocumentSource) {
      options.getRef()?.load(document);
    },

    async save(): Promise<Uint8Array | null> {
      const buf = await options.getRef()?.save();
      return buf ? new Uint8Array(buf) : null;
    },

    focus() {
      options.getRef()?.focus();
    },

    insertTextAtCursor(
      text: string,
      dropPositionOrOptions?: CursorPosition | InsertTextOptions
    ): InsertTextResult {
      const ed = options.getRef()?.getEditor() ?? null;
      if (!ed) {
        return {
          ok: false,
          changed: false,
          mode: 'failed',
          fallbackApplied: false,
          reason: 'editor-not-ready',
        };
      }

      const isRawPosition = (
        v: CursorPosition | InsertTextOptions | undefined
      ): v is CursorPosition =>
        Boolean(v && typeof v === 'object' && 'blockId' in v && 'offset' in v);

      const dropPosition = isRawPosition(dropPositionOrOptions)
        ? dropPositionOrOptions
        : dropPositionOrOptions?.dropPosition;
      const allowFallback = isRawPosition(dropPositionOrOptions)
        ? false
        : Boolean(dropPositionOrOptions?.allowFallback);

      const readCaret = (): CursorPosition | undefined => {
        try {
          const head = (
            ed as Editor & { getSelectionRange?: () => { head?: CursorPosition } | null }
          ).getSelectionRange?.()?.head;
          return head ? { blockId: head.blockId, offset: head.offset } : undefined;
        } catch {
          return undefined;
        }
      };

      try {
        ed.focus();
        if (dropPosition) {
          const sel = ed.exec({
            type: 'setSelection',
            range: { anchor: dropPosition, head: dropPosition },
          });
          if (!sel.ok) {
            return {
              ok: false,
              changed: false,
              mode: 'failed',
              fallbackApplied: false,
              reason: 'invalid-drop-position',
            };
          }
        }

        const beforeCaret = dropPosition ?? readCaret();
        let res = ed.exec({ type: 'replaceText', text });
        let fallbackApplied = false;
        let landedPosition: CursorPosition | undefined = readCaret() ?? beforeCaret;

        if ((!res.ok || !res.changed) && !dropPosition && allowFallback) {
          const paragraphs = ed.getParagraphs();
          const target = paragraphs.find((p) => p.text.trim().length > 0) ?? paragraphs[0];
          if (target) {
            const caret = { blockId: target.blockId, offset: target.text.length };
            const moved = ed.exec({
              type: 'setSelection',
              range: { anchor: caret, head: caret },
            });
            if (moved.ok) {
              res = ed.exec({ type: 'replaceText', text });
              fallbackApplied = true;
              landedPosition = readCaret() ?? caret;
            }
          }
        }

        const changed = res.ok ? Boolean(res.changed) : false;
        const mode: InsertTextMode =
          res.ok && changed
            ? fallbackApplied
              ? 'fallback-first-paragraph'
              : dropPosition
                ? 'drop-position'
                : 'cursor'
            : 'failed';

        return {
          ok: Boolean(res.ok),
          changed,
          mode,
          fallbackApplied,
          position: res.ok && changed ? landedPosition : undefined,
          reason: res.ok
            ? changed
              ? undefined
              : 'no-change'
            : res.reason || 'no-active-selection',
        };
      } catch (error) {
        return {
          ok: false,
          changed: false,
          mode: 'failed',
          fallbackApplied: false,
          reason: error instanceof Error ? error.message : String(error),
        };
      }
    },

    getPageCount,

    getDiagnostics(): WebOfficeDiagnostics {
      return {
        engine: 'eigenpal-apache-core',
        pageCount: getPageCount(),
        browser: probeBrowserEnvironment(),
        fonts: options.getFontDiagnostics?.(),
      };
    },

    getRawEditor() {
      return options.getRef()?.getEditor() ?? null;
    },

    getRawRef() {
      return options.getRef();
    },
  };
}
