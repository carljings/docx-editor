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
