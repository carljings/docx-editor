import type { OoxmlElement, OoxmlPart } from '../store/package/ooxml-tree.ts';
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

/** Validate optional `shading`, `borders`, and `borderMode` arguments on `setParagraphFormat`. */
export function paragraphVisualArgRefusal(
  shading: unknown,
  borders: unknown,
  borderMode?: unknown
): string | null {
  if (borderMode !== undefined && borderMode !== 'replace' && borderMode !== 'merge') {
    return "borderMode must be 'replace' or 'merge'";
  }
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

function isHeaderRowNode(row: OoxmlElement): boolean {
  const trPr = wmlChildNamed(row, 'trPr');
  const tblHeader = trPr ? wmlChildNamed(trPr, 'tblHeader') : undefined;
  if (!tblHeader) return false;
  const val = wmlAttributeValue(tblHeader, 'val')?.toLowerCase();
  return val !== '0' && val !== 'false' && val !== 'off';
}

/** Count contiguous leading rows marked `w:tblHeader` (ECMA-376 §17.4.50). */
export function tableHeaderRowCount(part: OoxmlPart, tableId: string): number {
  const topology = readEditableTableTopology(part.root, tableId);
  if (!topology.ok) return 0;
  let count = 0;
  for (const entry of topology.topology.rows) {
    if (!isHeaderRowNode(entry.row)) break;
    count += 1;
  }
  return count;
}

/**
 * Build a `setTableProperties` op toggling repeating header row(s) for the caret's row.
 *
 * OOXML `w:tblHeader` applies only to a contiguous prefix starting at row 0 (§17.4.50).
 * Non-adjacent body rows (`rowIndex > current`) return `null` so an off-state click on a
 * body row never wipes out existing header rows.
 */
export function toggleHeaderRowOp(
  part: OoxmlPart,
  tableId: string,
  rowId?: string
): TreeDocOp | null {
  const topology = readEditableTableTopology(part.root, tableId);
  if (!topology.ok || topology.topology.rows.length === 0) return null;
  const rows = topology.topology.rows;
  const rowIndex = rowId ? rows.findIndex((entry) => entry.row.id === rowId) : 0;
  if (rowIndex < 0) return null;
  const current = tableHeaderRowCount(part, tableId);
  let nextHeaderRowCount: number;
  if (rowIndex < current) {
    nextHeaderRowCount = rowIndex;
  } else if (rowIndex === current && current < rows.length) {
    nextHeaderRowCount = current + 1;
  } else {
    return null;
  }
  return {
    op: 'setTableProperties',
    tableId,
    headerRowCount: nextHeaderRowCount,
  };
}
