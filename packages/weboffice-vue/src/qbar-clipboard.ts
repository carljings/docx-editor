import type {
  Editor,
  EditorSelectionRange,
  SupportedImageMime,
} from '@docx-editor.dev/core/contracts/editor';

export interface CutPrecondition {
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

export function captureCutPrecondition(
  ed: Editor,
  drawingNodeId: string | null
): CutPrecondition | null {
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

export function matchesCutPrecondition(
  before: CutPrecondition,
  currentEditor: Editor | null
): boolean {
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

export function isSupportedRasterMime(mime: string | null): mime is SupportedImageMime {
  return mime === 'image/png' || mime === 'image/jpeg' || mime === 'image/gif';
}

function isClipboardPermissionError(err: unknown): boolean {
  const name = (err as { name?: unknown } | null)?.name;
  return name === 'NotAllowedError' || name === 'SecurityError';
}

export async function writeClipboardRichChecked(
  text: string,
  html: string | null
): Promise<boolean> {
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

export async function readClipboardPayload(
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

export async function executeClipboardCopyOrCut(opts: {
  readonly ed: Editor;
  readonly getCurrentEditor: () => Editor | null;
  readonly action: 'copy' | 'cut';
  readonly cutTicket: number;
  readonly getLatestCutTicket: () => number;
}): Promise<string> {
  const { ed, getCurrentEditor, action, cutTicket, getLatestCutTicket } = opts;
  const gate = ed.can({ type: action });
  if (!gate.ok) {
    return action === 'copy' ? '请先选中要复制的内容' : '请先选中要剪切的内容';
  }
  const flavours = ed.query({ type: 'copyFlavours' });
  if (!flavours || flavours.reason || (!flavours.text && !flavours.html)) {
    return action === 'copy' ? '当前选区不支持复制' : '当前选区不支持剪切';
  }
  const cutPrecondition =
    action === 'cut' ? captureCutPrecondition(ed, flavours.drawingNodeId ?? null) : null;
  if (action === 'cut' && !cutPrecondition) {
    return '当前文档状态不稳定，请重新剪切';
  }
  const wrote = await writeClipboardRichChecked(flavours.text, flavours.html);
  if (!wrote) {
    if (getCurrentEditor() === ed) ed.focus();
    return action === 'copy'
      ? '剪贴板写入被拒绝，请使用 Ctrl+C 复制'
      : '剪贴板写入被拒绝，未删除所选内容（请使用 Ctrl+X）';
  }
  if (action === 'cut') {
    if (
      cutTicket !== getLatestCutTicket() ||
      !cutPrecondition ||
      !matchesCutPrecondition(cutPrecondition, getCurrentEditor())
    ) {
      if (getCurrentEditor() === ed) ed.focus();
      return '选区或文档已变化，内容已复制但未删除';
    }
    const delRes = ed.exec({ type: 'deleteText' });
    if (!delRes.ok) {
      if (getCurrentEditor() === ed) ed.focus();
      return `剪切失败：${delRes.reason || '当前位置不可编辑'}`;
    }
  }
  if (getCurrentEditor() === ed) ed.focus();
  return action === 'copy' ? '已复制' : '已剪切';
}

export async function executeClipboardPaste(
  ed: Editor,
  plainOnly: boolean
): Promise<string | null> {
  if (!navigator.clipboard) {
    ed.focus();
    return '请直接使用快捷键 Ctrl+V / Ctrl+Shift+V 粘贴';
  }
  try {
    const { text, html } = await readClipboardPayload(plainOnly);
    if (!text && !html) {
      ed.focus();
      return '剪贴板为空';
    }
    const res = plainOnly
      ? ed.exec({ type: 'pasteWithoutFormatting', text })
      : ed.exec({ type: 'paste', text, ...(html ? { html } : {}) });
    ed.focus();
    return res.ok ? null : `粘贴失败：${res.reason || '当前位置不可编辑'}`;
  } catch {
    ed.focus();
    return '浏览器限制读取剪贴板，请按 Ctrl+V / Ctrl+Shift+V';
  }
}
