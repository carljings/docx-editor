// The text ops that replace one paragraph range with new text, shared by every replacing
// command: `replaceMatch`, `replaceAllMatches`, and `replaceAcceptedText`.
//
// Deleting first and then inserting at the range start put the new text wherever an insert at
// that boundary goes, which is the run BEFORE the range, or the run after a hyperlink when
// the range ended one. A fix of the last word in a link moved its letters out of the link. So
// an untracked replacement inserts first, at the range start with a right bias, which joins
// the run of the first replaced character, and then deletes the old text after it. A tracked
// replacement keeps the strike and inserts where the struck text ends, as typing does.

import { findNode, type OoxmlPart, type TreeDocOp } from '@docx-editor.dev/core/store';
import { isContentRevisionKind } from '../store/package/ooxml-shared.ts';
import {
  inlineContainersOf,
  isShowingPlaceholder,
  parentOf,
} from '../store/store/tree-op-nodes.ts';
import { segmentsOf } from '../store/store/tree-op-segments.ts';
import { textFormFieldsOf } from '../store/store/text-form-fields.ts';
import { contentControlAtCaret } from '../store/store/tree-op-validate-controls.ts';

/**
 * Ops that replace `[start, end)` of a paragraph with `text`. `landing` is where a tracked
 * replacement inserts, or null when the edit is not tracked.
 */
export function replacementOps(
  paragraphId: string,
  start: number,
  end: number,
  text: string,
  landing: number | null
): TreeDocOp[] {
  if (end <= start) {
    return text.length > 0
      ? [{ op: 'insertText', paragraphId, offset: landing ?? start, text }]
      : [];
  }
  if (text.length === 0) return [{ op: 'deleteText', paragraphId, start, end }];
  if (landing !== null) {
    return [
      { op: 'deleteText', paragraphId, start, end },
      { op: 'insertText', paragraphId, offset: landing, text },
    ];
  }
  return [
    { op: 'insertText', paragraphId, offset: start, text, bias: 'right' },
    {
      op: 'deleteText',
      paragraphId,
      start: start + text.length,
      end: end + text.length,
    },
  ];
}

function canInsertIntoReplacedRun(part: OoxmlPart, paragraphId: string, offset: number): boolean {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return false;
  const control = contentControlAtCaret(part, paragraph, offset, offset, 'right');
  if (control && isShowingPlaceholder(control)) return false;
  if (textFormFieldsOf(paragraph).some((field) => offset >= field.start && offset < field.end)) {
    return false;
  }
  const segment = segmentsOf(paragraph).find((item) => item.start <= offset && offset < item.end);
  if (!segment || segment.removeNodeIds !== undefined || segment.node.kind !== 'textValue') {
    return false;
  }
  const textParent = parentOf(part, segment.node.id);
  if (!textParent || textParent.kind !== 'text') return false;
  const containers = inlineContainersOf(paragraph, segment.runId);
  if (
    containers.some(
      (container) => isContentRevisionKind(container.kind) || isShowingPlaceholder(container)
    )
  ) {
    return false;
  }
  return true;
}

/**
 * Combine a selection deletion plan with an insertion, preserving the first replaced
 * character's run (and its full `w:rPr`) when an untracked deletion would otherwise sweep
 * an emptied run before `insertText` executes.
 */
export function selectionReplacementInsertOps(
  part: OoxmlPart,
  planOps: readonly TreeDocOp[],
  insertion: TreeDocOp,
  options?: { readonly outsidePrompt?: boolean }
): TreeDocOp[] {
  if (
    options?.outsidePrompt ||
    insertion.op !== 'insertText' ||
    insertion.text.length === 0 ||
    insertion.revision !== undefined ||
    insertion.inside !== undefined ||
    planOps.length === 0
  ) {
    return [
      ...planOps,
      options?.outsidePrompt && insertion.op === 'insertText'
        ? { ...insertion, bias: 'left' }
        : insertion,
    ];
  }
  const { paragraphId, offset, text } = insertion;
  let deleteIndex = -1;
  for (let index = 0; index < planOps.length; index += 1) {
    const op = planOps[index]!;
    if (op.op === 'deleteText' && op.paragraphId === paragraphId) {
      if (
        deleteIndex !== -1 ||
        op.revision !== undefined ||
        op.start !== offset ||
        op.end <= offset
      ) {
        deleteIndex = -1;
        break;
      }
      deleteIndex = index;
      continue;
    }
    if (
      deleteIndex === -1 &&
      ((op.op === 'joinParagraphs' &&
        (op.firstId === paragraphId || op.secondId === paragraphId)) ||
        (op.op === 'deleteBlock' && op.blockId === paragraphId))
    ) {
      deleteIndex = -1;
      break;
    }
  }
  if (deleteIndex === -1 || !canInsertIntoReplacedRun(part, paragraphId, offset)) {
    return [...planOps, insertion];
  }
  const deleteOp = planOps[deleteIndex] as Extract<TreeDocOp, { op: 'deleteText' }>;
  const shiftedDelete: TreeDocOp = {
    ...deleteOp,
    start: deleteOp.start + text.length,
    end: deleteOp.end + text.length,
  };
  return [
    { ...insertion, bias: 'right' },
    ...planOps.slice(0, deleteIndex),
    shiftedDelete,
    ...planOps.slice(deleteIndex + 1),
  ];
}
