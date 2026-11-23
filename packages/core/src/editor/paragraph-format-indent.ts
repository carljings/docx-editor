import type { OoxmlElement, OoxmlPart } from '../store/package/ooxml-tree.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import { validateTreeOp, type TreeDocOp } from '../store/store/tree-ops.ts';
import type { TableBorderEdgeTarget, TableBorderSpecInput } from '../store/store/tree-op-types.ts';

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

export type WholeTableBorderPreset = 'threeLine' | 'tableAll' | 'tableNone';

export function isWholeTableBorderPreset(scope: unknown): scope is WholeTableBorderPreset {
  return scope === 'threeLine' || scope === 'tableAll' || scope === 'tableNone';
}

const DEFAULT_THREE_LINE_SPEC: TableBorderSpecInput = {
  style: 'single',
  size: 12,
  color: { kind: 'hex', value: '000000' },
};

const DEFAULT_TABLE_ALL_SPEC: TableBorderSpecInput = {
  style: 'single',
  size: 4,
  color: { kind: 'hex', value: '334155' },
};

/**
 * Build atomic whole-table `setTableCellBorders` ops for official table presets:
 * - `'threeLine'`: GB/T 7713 / official three-line table (1.5pt top & bottom, 0.75pt under header).
 * - `'tableAll'`: uniform thin grid across the whole table.
 * - `'tableNone'`: explicit `val="none"` across all cells (overriding `TableGrid` style).
 */
export function wholeTableBorderPresetOps(
  part: OoxmlPart,
  tableId: string,
  preset: WholeTableBorderPreset,
  spec?: TableBorderSpecInput
): TreeDocOp[] | null {
  const topology = readEditableTableTopology(part.root, tableId);
  if (!topology.ok || topology.topology.rows.length === 0) return null;
  const rows = topology.topology.rows;
  const allCellIds = [...new Set(rows.flatMap((row) => row.cells.map((cell) => cell.id)))];
  if (allCellIds.length === 0) return null;

  if (preset === 'tableNone') {
    return [
      { op: 'setTableCellBorders', tableId, cellIds: allCellIds, scope: 'none', target: 'all' },
    ];
  }
  if (preset === 'tableAll') {
    return [
      {
        op: 'setTableCellBorders',
        tableId,
        cellIds: allCellIds,
        scope: 'all',
        spec: spec ?? DEFAULT_TABLE_ALL_SPEC,
      },
    ];
  }

  const thick = spec ?? DEFAULT_THREE_LINE_SPEC;
  const thin: TableBorderSpecInput = {
    style: thick.style,
    size: Math.max(4, Math.round(thick.size / 2)),
    color: thick.color,
  };
  const ops: TreeDocOp[] = [
    { op: 'setTableCellBorders', tableId, cellIds: allCellIds, scope: 'none', target: 'all' },
    { op: 'setTableCellBorders', tableId, cellIds: allCellIds, scope: 'top', spec: thick },
    { op: 'setTableCellBorders', tableId, cellIds: allCellIds, scope: 'bottom', spec: thick },
  ];

  if (rows.length > 1) {
    const headerCount = tableHeaderRowCount(part, tableId);
    const startRow = headerCount > 0 && headerCount < rows.length ? headerCount - 1 : 0;
    for (let endRow = startRow; endRow < rows.length - 1; endRow += 1) {
      const headerCellIds = [
        ...new Set(rows.slice(0, endRow + 1).flatMap((row) => row.cells.map((cell) => cell.id))),
      ];
      const midOp: TreeDocOp = {
        op: 'setTableCellBorders',
        tableId,
        cellIds: headerCellIds,
        scope: 'bottom',
        spec: thin,
      };
      if (validateTreeOp(part, midOp) === null) {
        ops.push(midOp);
        const bodyCellIds = [
          ...new Set(rows.slice(endRow + 1).flatMap((row) => row.cells.map((cell) => cell.id))),
        ];
        if (bodyCellIds.length > 0) {
          const bodyTopOp: TreeDocOp = {
            op: 'setTableCellBorders',
            tableId,
            cellIds: bodyCellIds,
            scope: 'top',
            spec: thin,
          };
          if (validateTreeOp(part, bodyTopOp) === null) ops.push(bodyTopOp);
        }
        break;
      }
    }
  }
  return ops;
}

export function planSetTableBordersOps(args: {
  readonly part: OoxmlPart;
  readonly tableId: string;
  readonly cellIds: readonly string[];
  readonly scope: string;
  readonly target?: string;
  readonly spec?: unknown;
  readonly rtl: boolean;
  readonly lowerSpec: (
    spec: unknown
  ) =>
    | { readonly ok: true; readonly spec: TableBorderSpecInput }
    | { readonly ok: false; readonly reason: string };
}):
  | { readonly ok: true; readonly ops: readonly TreeDocOp[] }
  | { readonly ok: false; readonly code: 'invalidArgs' | 'unsupported'; readonly reason: string } {
  const { part, tableId, cellIds, scope, target, spec, rtl, lowerSpec } = args;
  const flip = (edge: TableBorderEdgeTarget): TableBorderEdgeTarget =>
    rtl ? (edge === 'left' ? 'right' : edge === 'right' ? 'left' : edge) : edge;
  if (isWholeTableBorderPreset(scope)) {
    const lowered = spec !== undefined ? lowerSpec(spec) : null;
    if (lowered && !lowered.ok) return { ok: false, code: 'invalidArgs', reason: lowered.reason };
    const ops = wholeTableBorderPresetOps(part, tableId, scope, lowered?.spec);
    return ops
      ? { ok: true, ops }
      : { ok: false, code: 'unsupported', reason: 'cannot apply whole-table border preset' };
  }
  if (scope === 'none') {
    const op: TreeDocOp = {
      op: 'setTableCellBorders',
      tableId,
      cellIds,
      scope: 'none',
      target: flip(target as TableBorderEdgeTarget),
    };
    return { ok: true, ops: [op] };
  }
  const lowered = lowerSpec(spec);
  if (!lowered.ok) return { ok: false, code: 'invalidArgs', reason: lowered.reason };
  const op: TreeDocOp = {
    op: 'setTableCellBorders',
    tableId,
    cellIds,
    scope: flip(scope as TableBorderEdgeTarget),
    spec: lowered.spec,
  };
  return { ok: true, ops: [op] };
}
