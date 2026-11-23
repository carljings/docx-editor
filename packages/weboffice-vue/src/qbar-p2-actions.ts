import type { Editor } from '@docx-editor.dev/core/contracts/editor';
import {
  DEFAULT_IMAGE_RESOURCE_LIMITS,
  sniffImageMime as sniffImageBytesMime,
  validateRasterHeader,
} from '@docx-editor.dev/core/editor';
import { isSupportedRasterMime } from './qbar-clipboard';
import type { HeaderFooterQuickAction } from './qbar-p2-menus';

export type TableContextQuickAction =
  | 'mergeCells'
  | 'toggleHeaderRow'
  | 'selectTable'
  | 'rowAbove'
  | 'rowBelow'
  | 'colLeft'
  | 'colRight'
  | 'deleteRow'
  | 'deleteColumn'
  | 'deleteTable';

export function runTableQuickAction(
  ed: Editor,
  action: TableContextQuickAction
): { readonly ok: boolean; readonly message?: string } {
  const res =
    action === 'mergeCells'
      ? ed.exec({ type: 'mergeCells' })
      : action === 'toggleHeaderRow'
        ? ed.exec({ type: 'toggleHeaderRow' })
        : action === 'selectTable'
          ? ed.exec({ type: 'selectTableRegion', region: 'table' })
          : action === 'rowAbove'
            ? ed.exec({ type: 'insertRow', where: 'above' })
            : action === 'rowBelow'
              ? ed.exec({ type: 'insertRow', where: 'below' })
              : action === 'colLeft'
                ? ed.exec({ type: 'insertColumn', where: 'left' })
                : action === 'colRight'
                  ? ed.exec({ type: 'insertColumn', where: 'right' })
                  : action === 'deleteRow'
                    ? ed.exec({ type: 'deleteRow' })
                    : action === 'deleteColumn'
                      ? ed.exec({ type: 'deleteColumn' })
                      : ed.exec({ type: 'deleteTable' });
  if (!res.ok) return { ok: false, message: res.reason || '当前表格位置无法执行此操作' };
  if (action === 'toggleHeaderRow') {
    const nextHeaders = ed.snapshot().table?.headerRowCount ?? 0;
    return {
      ok: true,
      message:
        nextHeaders === 0
          ? '已取消重复表头行'
          : nextHeaders === 1
            ? '已开启首行重复表头'
            : `已设前 ${nextHeaders} 行为重复表头`,
    };
  }
  if (action === 'mergeCells') return { ok: true, message: '已合并选中单元格' };
  if (action === 'selectTable') return { ok: true, message: '已选中整张表格' };
  return { ok: true };
}

export async function insertRasterImageFromFile(
  ed: Editor,
  file: File
): Promise<{ readonly ok: boolean; readonly message: string }> {
  try {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);
    const limits = DEFAULT_IMAGE_RESOURCE_LIMITS;
    if (bytes.byteLength === 0) return { ok: false, message: '图片文件为空' };
    if (bytes.byteLength > limits.maxEncodedBytes)
      return { ok: false, message: '图片体积超出限制' };
    const sniffed = sniffImageBytesMime(bytes);
    if (!isSupportedRasterMime(sniffed))
      return { ok: false, message: '仅支持 PNG / JPG / GIF 图片' };
    const header = validateRasterHeader(bytes, sniffed);
    if (!header) return { ok: false, message: '图片文件头无效或已损坏' };
    if (
      header.pixelWidth * header.pixelHeight > limits.maxPixels ||
      header.pixelWidth > limits.maxDimension ||
      header.pixelHeight > limits.maxDimension
    ) {
      return { ok: false, message: '图片分辨率超出限制' };
    }
    const dpiX = header.dpiX ?? 96;
    const dpiY = header.dpiY ?? 96;
    const widthPoints = Math.max(1, Math.round(((header.pixelWidth * 72) / dpiX) * 100) / 100);
    const heightPoints = Math.max(1, Math.round(((header.pixelHeight * 72) / dpiY) * 100) / 100);
    const res = await ed.executeImageCommand({
      type: 'insertImage',
      data: bytes,
      mime: sniffed,
      widthPoints,
      heightPoints,
    });
    return res.ok
      ? { ok: true, message: '已插入图片' }
      : { ok: false, message: `插入图片失败：${res.reason || '当前位置不支持'}` };
  } catch {
    return { ok: false, message: '读取图片失败' };
  }
}

function prepareCenteredFooterParagraph(ed: Editor): void {
  const range = ed.getSelectionRange();
  if (!range) return;
  const para = ed.getParagraphText(range.head.blockId);
  if (para && para.text.trim().length > 0) {
    const endPos = { blockId: para.blockId, offset: para.text.length };
    ed.exec({
      type: 'setSelection',
      range: { anchor: endPos, head: endPos },
    });
    ed.exec({ type: 'splitParagraph' });
  }
}

/**
 * Execute header/footer/page-number actions cleanly against Core's OOXML commands.
 * Resolves the active caret section even from body scope, protects existing headers
 * when inserting official centered `— X —` footer page numbers, and preserves any
 * pre-existing footer text on its own line.
 */
export function runHeaderFooterQuickAction(
  ed: Editor,
  action: HeaderFooterQuickAction
): { readonly ok: boolean; readonly message: string; readonly titlePage?: boolean } {
  const hfState = ed.getHeaderFooterState();
  const sf = ed.query({ type: 'sectionFurniture' });
  const targetSectionIndex = hfState?.sectionIndex ?? sf?.sectionIndex ?? 0;
  const wasEditing = hfState?.editing ?? null;

  if (action === 'editHeader' || action === 'editFooter') {
    const position = action === 'editHeader' ? 'header' : 'footer';
    const res = ed.exec({ type: 'editHeaderFooter', position, sectionIndex: targetSectionIndex });
    return res.ok
      ? { ok: true, message: position === 'header' ? '已进入页眉编辑' : '已进入页脚编辑' }
      : { ok: false, message: res.reason || '当前位置无法编辑页眉页脚' };
  }

  if (action === 'exitHeaderFooter') {
    const res = ed.exec({ type: 'exitHeaderFooter' });
    return res.ok
      ? { ok: true, message: '已返回正文编辑' }
      : { ok: false, message: res.reason || '无法退出页眉页脚' };
  }

  if (action === 'toggleFirstPage') {
    const current = hfState?.titlePage ?? sf?.titlePage ?? false;
    const next = !current;
    const res = ed.exec({
      type: 'setHeaderFooterOptions',
      sectionIndex: targetSectionIndex,
      titlePage: next,
    });
    return res.ok
      ? { ok: true, message: next ? '已开启首页不同' : '已关闭首页不同', titlePage: next }
      : { ok: false, message: res.reason || '当前无法切换首页不同' };
  }

  if (action === 'removeHeader' || action === 'removeFooter') {
    const position = action === 'removeHeader' ? 'header' : 'footer';
    const res = ed.exec({ type: 'removeHeaderFooter', position, sectionIndex: targetSectionIndex });
    return res.ok
      ? { ok: true, message: position === 'header' ? '已删除页眉' : '已删除页脚' }
      : { ok: false, message: res.reason || '当前节没有可删除的页眉或页脚' };
  }

  if (action === 'insertOfficialPageNum') {
    // Official folio (`— X —`) ALWAYS targets the active section's footer, never polluting headers.
    const mustSwitchToFooter = wasEditing !== 'footer';
    if (wasEditing === 'header') {
      ed.exec({ type: 'exitHeaderFooter' });
    }
    if (mustSwitchToFooter) {
      const entered = ed.exec({
        type: 'editHeaderFooter',
        position: 'footer',
        sectionIndex: targetSectionIndex,
      });
      if (!entered.ok) {
        return { ok: false, message: entered.reason || '当前文档无法插入页码' };
      }
      prepareCenteredFooterParagraph(ed);
    }

    const alignRes = ed.exec({ type: 'setAlignment', align: 'center' });
    if (!alignRes.ok) {
      if (mustSwitchToFooter) ed.exec({ type: 'exitHeaderFooter' });
      return { ok: false, message: alignRes.reason || '无法设置页码居中' };
    }
    const leftRes = ed.exec({ type: 'replaceText', text: '— ' });
    if (!leftRes.ok) {
      if (mustSwitchToFooter) ed.exec({ type: 'exitHeaderFooter' });
      return { ok: false, message: leftRes.reason || '无法插入页码修饰符' };
    }
    const fieldRes = ed.exec({ type: 'insertPageField', field: 'PAGE' });
    if (!fieldRes.ok) {
      if (mustSwitchToFooter) ed.exec({ type: 'exitHeaderFooter' });
      return { ok: false, message: fieldRes.reason || '无法插入页码域' };
    }
    ed.exec({ type: 'replaceText', text: ' —' });
    if (mustSwitchToFooter) ed.exec({ type: 'exitHeaderFooter' });
    return { ok: true, message: '已在页脚插入公文页码（— X —）' };
  }

  // Generic page fields insert at the caret when already editing a header/footer,
  // or auto-enter the current section's footer when triggered from body scope.
  let enteredTemporarily = false;
  if (wasEditing === null) {
    const entered = ed.exec({
      type: 'editHeaderFooter',
      position: 'footer',
      sectionIndex: targetSectionIndex,
    });
    if (!entered.ok) {
      return { ok: false, message: entered.reason || '当前文档无法插入页码' };
    }
    enteredTemporarily = true;
    prepareCenteredFooterParagraph(ed);
    ed.exec({ type: 'setAlignment', align: 'center' });
  }

  const field =
    action === 'insertPage'
      ? ('PAGE' as const)
      : action === 'insertNumPages'
        ? ('NUMPAGES' as const)
        : ('PAGE_X_OF_Y' as const);
  const res = ed.exec({ type: 'insertPageField', field });
  if (enteredTemporarily) ed.exec({ type: 'exitHeaderFooter' });
  if (!res.ok) return { ok: false, message: res.reason || '无法插入页码域' };
  const label =
    field === 'PAGE' ? '当前页码' : field === 'NUMPAGES' ? '总页数' : '第 X 页 / 共 Y 页';
  return { ok: true, message: `已插入${label}` };
}
