import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import type { TreeDocOp } from '../store/store/tree-ops.ts';

/**
 * Write signed character-unit first-line indent (`w:firstLineChars` / `w:hangingChars`)
 * onto a `w:ind` attribute bag alongside an optional twip fallback (§17.3.1.12).
 */
export function writeFirstLineChars(
  authored: Readonly<Record<string, string>>,
  chars: number | null,
  fallbackTwips: number | null | undefined
): Record<string, string> {
  const attributes: Record<string, string> = { ...authored };
  delete attributes.firstLineChars;
  delete attributes.hangingChars;
  if (chars === null) {
    delete attributes.firstLine;
    delete attributes.hanging;
    return attributes;
  }
  if (chars === 0) {
    attributes.firstLine = '0';
    attributes.firstLineChars = '0';
    attributes.hanging = '0';
    attributes.hangingChars = '0';
    return attributes;
  }
  if (chars > 0) {
    attributes.firstLineChars = String(chars);
    attributes.hanging = '0';
    if (fallbackTwips !== undefined && fallbackTwips !== null) {
      attributes.firstLine = String(Math.max(0, fallbackTwips));
    } else {
      delete attributes.firstLine;
    }
  } else {
    attributes.hangingChars = String(-chars);
    attributes.firstLine = '0';
    if (fallbackTwips !== undefined && fallbackTwips !== null) {
      attributes.hanging = String(Math.max(0, -fallbackTwips));
    } else {
      delete attributes.hanging;
    }
  }
  return attributes;
}

/** Build `setListNumbering` ops for an explicit Chinese official list preset or clear. */
export function buildPresetListOps(
  touched: readonly string[],
  preset: 'cn-heading1' | 'cn-heading2' | 'cn-decimal' | 'cn-paren' | 'clear',
  deps: {
    readonly ensureListDefinition: (
      kind: 'ordered',
      geometry: undefined,
      preset: 'cn-heading1' | 'cn-heading2' | 'cn-decimal' | 'cn-paren'
    ) => string | null;
    readonly leaveOp: (
      paragraphId: string
    ) => import('@docx-editor.dev/core/store').TreeDocOp | null;
  }
): import('@docx-editor.dev/core/store').TreeDocOp[] | null {
  if (preset === 'clear') {
    const ops: import('@docx-editor.dev/core/store').TreeDocOp[] = [];
    for (const paragraphId of touched) {
      const op = deps.leaveOp(paragraphId);
      if (op) ops.push(op);
    }
    return ops.length > 0 ? ops : null;
  }
  const numId = deps.ensureListDefinition('ordered', undefined, preset);
  if (!numId) return null;
  return touched.map((paragraphId) => ({
    op: 'setListNumbering' as const,
    paragraphId,
    numId,
    level: 0,
  }));
}

export const UNDERLINE_VARIANTS: ReadonlySet<string> = new Set([
  'single',
  'double',
  'thick',
  'dotted',
  'dash',
  'dotDash',
  'wave',
  'none',
]);

const BORDER_EDGE_NAMES: ReadonlySet<string> = new Set([
  'top',
  'left',
  'bottom',
  'right',
  'between',
  'bar',
]);

const HEX_6 = /^[0-9a-fA-F]{6}$/;

/** Validate optional `shading` and `borders` arguments on `setParagraphFormat`. */
export function paragraphVisualArgRefusal(shading: unknown, borders: unknown): string | null {
  if (shading !== undefined && shading !== null) {
    if (typeof shading !== 'string' || !HEX_6.test(shading)) {
      return 'shading requires a six-digit hex RGB value like F1F5F9 or null';
    }
  }
  if (borders !== undefined) {
    if (!Array.isArray(borders)) return 'borders must be an array of edge definitions';
    for (const entry of borders) {
      if (
        typeof entry !== 'object' ||
        entry === null ||
        typeof (entry as { edge?: unknown }).edge !== 'string' ||
        !BORDER_EDGE_NAMES.has((entry as { edge: string }).edge) ||
        typeof (entry as { attributes?: unknown }).attributes !== 'object' ||
        (entry as { attributes?: unknown }).attributes === null ||
        typeof (entry as { attributes: { val?: unknown } }).attributes.val !== 'string'
      ) {
        return 'each border edge requires a valid edge name and attributes.val';
      }
    }
  }
  return null;
}

/** Build a `setTableProperties` op toggling the table's repeating header row on or off. */
export function toggleHeaderRowOp(part: OoxmlPart, tableId: string): TreeDocOp | null {
  const topology = readEditableTableTopology(part.root, tableId);
  if (!topology.ok || topology.topology.rows.length === 0) return null;
  const firstRow = topology.topology.rows[0]!.row;
  const trPr = wmlChildNamed(firstRow, 'trPr');
  const tblHeader = trPr ? wmlChildNamed(trPr, 'tblHeader') : null;
  const val = tblHeader ? wmlAttributeValue(tblHeader, 'val')?.toLowerCase() : null;
  const isHeader = tblHeader !== null && val !== '0' && val !== 'false' && val !== 'off';
  return {
    op: 'setTableProperties',
    tableId,
    headerRowCount: isHeader ? 0 : 1,
  };
}
