// P1 official-document helpers for preset numbering continuity and paragraph border merging.

import {
  findNode,
  parentNodeOf,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
  type OrderedListPreset,
  type ParagraphBorderEdgeName,
  type ParagraphBorderEdgeWrite,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import type { ListMarkerRecord } from '@docx-editor.dev/core/layout';
import { wmlAttributeValue, wmlChildNamed } from '../store/store/tree-op-table-shared.ts';

const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICIAL_PRESET_ORDER: readonly OrderedListPreset[] = [
  'cn-heading1',
  'cn-heading2',
  'cn-decimal',
  'cn-paren',
];
const MAX_LIST_LEVEL = 8;

interface ResolvedLevelFmt {
  readonly numFmt: string;
  readonly lvlText: string;
}

function wmlChildrenNamed(parent: OoxmlElement, localName: string): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const child of parent.children) {
    if (
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NS &&
      child.localName === localName
    ) {
      out.push(child);
    }
  }
  return out;
}

function resolveNumberingLevelFmt(
  numberingPart: OoxmlPart | undefined,
  numId: string,
  level: number
): ResolvedLevelFmt | null {
  if (!numberingPart) return null;
  const root = numberingPart.root;
  const ilvlStr = String(level);
  const numNode = wmlChildrenNamed(root, 'num').find(
    (node) => wmlAttributeValue(node, 'numId') === numId
  );
  if (!numNode) return null;

  let lvlNode: OoxmlElement | undefined;
  for (const override of wmlChildrenNamed(numNode, 'lvlOverride')) {
    if (wmlAttributeValue(override, 'ilvl') !== ilvlStr) continue;
    const inner = wmlChildNamed(override, 'lvl');
    if (inner) {
      lvlNode = inner;
      break;
    }
  }
  if (!lvlNode) {
    const absRef = wmlChildNamed(numNode, 'abstractNumId');
    const absId = absRef ? wmlAttributeValue(absRef, 'val') : undefined;
    if (!absId) return null;
    const abstractNode = wmlChildrenNamed(root, 'abstractNum').find(
      (node) => wmlAttributeValue(node, 'abstractNumId') === absId
    );
    if (!abstractNode) return null;
    lvlNode = wmlChildrenNamed(abstractNode, 'lvl').find(
      (node) => wmlAttributeValue(node, 'ilvl') === ilvlStr
    );
  }
  if (!lvlNode) return null;
  const numFmtNode = wmlChildNamed(lvlNode, 'numFmt');
  const lvlTextNode = wmlChildNamed(lvlNode, 'lvlText');
  const numFmt = numFmtNode ? wmlAttributeValue(numFmtNode, 'val') : undefined;
  const lvlText = lvlTextNode ? wmlAttributeValue(lvlTextNode, 'val') : undefined;
  if (!numFmt || lvlText === undefined) return null;
  return { numFmt, lvlText };
}

function levelMatchesPreset(
  numberingPart: OoxmlPart | undefined,
  numId: string,
  level: number,
  preset: OrderedListPreset
): boolean {
  const fmt = resolveNumberingLevelFmt(numberingPart, numId, level);
  if (!fmt) return false;
  const norm = fmt.lvlText.replace(/%[1-9]/g, '%N');
  switch (preset) {
    case 'cn-heading1':
      return (
        (fmt.numFmt === 'chineseCountingThousand' || fmt.numFmt === 'chineseCounting') &&
        norm === '%N、'
      );
    case 'cn-heading2':
      return (
        (fmt.numFmt === 'chineseCountingThousand' || fmt.numFmt === 'chineseCounting') &&
        (norm === '（%N）' || norm === '(%N)')
      );
    case 'cn-decimal':
      return fmt.numFmt === 'decimal' && norm === '%N.';
    case 'cn-paren':
      return fmt.numFmt === 'decimal' && (norm === '（%N）' || norm === '(%N)');
  }
}

function detectPresetOfMarker(
  numberingPart: OoxmlPart | undefined,
  marker: ListMarkerRecord | null
): OrderedListPreset | null {
  if (!marker) return null;
  for (const preset of OFFICIAL_PRESET_ORDER) {
    if (levelMatchesPreset(numberingPart, marker.numId, marker.level, preset)) return preset;
  }
  if (marker.level > 0) {
    for (const preset of OFFICIAL_PRESET_ORDER) {
      if (levelMatchesPreset(numberingPart, marker.numId, 0, preset)) return preset;
    }
  }
  return null;
}

function sameContainer(storyPart: OoxmlPart, a: string, b: string): boolean {
  return parentNodeOf(storyPart, a)?.id === parentNodeOf(storyPart, b)?.id;
}

export interface PresetListDeps {
  readonly numberingPart: OoxmlPart | undefined;
  readonly storyPart: OoxmlPart;
  readonly order: readonly string[];
  readonly markerOf: (paragraphId: string) => ListMarkerRecord | null;
  readonly ensureListDefinition: (
    kind: 'ordered',
    geometry: undefined,
    preset: OrderedListPreset
  ) => string | null;
  readonly leaveOp: (paragraphId: string) => TreeDocOp | null;
}

/**
 * Build `setListNumbering` ops for Chinese official numbering presets.
 *
 * - Reapplying the same preset preserves the current `numId` and `ilvl` (no-op).
 * - Applying beside or across a matching continuous list joins that `numId`.
 * - Switching to a different preset mints one `numId` and preserves relative sublevels.
 * - `restart: true` explicitly starts a fresh list instance from 1.
 */
export function buildPresetListOps(
  touched: readonly string[],
  preset: OrderedListPreset | 'clear' | undefined,
  restart: boolean,
  deps: PresetListDeps
): TreeDocOp[] | null {
  if (preset === 'clear') {
    const ops: TreeDocOp[] = [];
    for (const paragraphId of touched) {
      const op = deps.leaveOp(paragraphId);
      if (op) ops.push(op);
    }
    return ops;
  }

  const markers = touched.map((paragraphId) => deps.markerOf(paragraphId));
  const targetPreset: OrderedListPreset =
    preset ??
    markers
      .map((m) => detectPresetOfMarker(deps.numberingPart, m))
      .find((p): p is OrderedListPreset => p !== null) ??
    'cn-heading1';

  const matchesAt = (numId: string, level: number): boolean =>
    levelMatchesPreset(deps.numberingPart, numId, level, targetPreset);

  if (!restart) {
    // 1. Preserve current list instance & level when selection already matches the preset.
    if (touched.length === 1) {
      const m = markers[0];
      if (m) {
        if (matchesAt(m.numId, m.level)) return [];
        if (m.level > 0 && matchesAt(m.numId, 0)) {
          return [{ op: 'setListNumbering', paragraphId: touched[0]!, numId: m.numId, level: 0 }];
        }
      }
    } else if (
      markers.every(
        (m) => m !== null && (matchesAt(m.numId, m.level) || (m.level > 0 && matchesAt(m.numId, 0)))
      )
    ) {
      const firstNumId = markers[0]!.numId;
      if (markers.every((m) => m!.numId === firstNumId)) return [];
    }

    // 2. Reuse a matching continuous list inside the selection or immediately adjacent to it.
    let reusable: { readonly numId: string; readonly level: number } | null = null;
    for (const m of markers) {
      if (!m) continue;
      if (matchesAt(m.numId, m.level)) {
        reusable = { numId: m.numId, level: m.level };
        break;
      }
      if (m.level > 0 && matchesAt(m.numId, 0)) {
        reusable = { numId: m.numId, level: 0 };
        break;
      }
    }
    if (!reusable) {
      const firstIdx = deps.order.indexOf(touched[0]!);
      const lastIdx = deps.order.indexOf(touched[touched.length - 1]!);
      const candidates: Array<{ readonly neighborId: string; readonly edgeId: string }> = [];
      if (firstIdx > 0) {
        candidates.push({ neighborId: deps.order[firstIdx - 1]!, edgeId: touched[0]! });
      }
      if (lastIdx >= 0 && lastIdx + 1 < deps.order.length) {
        candidates.push({
          neighborId: deps.order[lastIdx + 1]!,
          edgeId: touched[touched.length - 1]!,
        });
      }
      for (const { neighborId, edgeId } of candidates) {
        if (!sameContainer(deps.storyPart, neighborId, edgeId)) continue;
        const nm = deps.markerOf(neighborId);
        if (!nm) continue;
        if (matchesAt(nm.numId, nm.level)) {
          reusable = { numId: nm.numId, level: nm.level };
          break;
        }
        if (nm.level > 0 && matchesAt(nm.numId, 0)) {
          reusable = { numId: nm.numId, level: 0 };
          break;
        }
      }
    }

    if (reusable) {
      const ops: TreeDocOp[] = [];
      for (let i = 0; i < touched.length; i += 1) {
        const paragraphId = touched[i]!;
        const m = markers[i];
        const keepSublevel =
          m !== null &&
          m.numId === reusable.numId &&
          (matchesAt(m.numId, m.level) ||
            (touched.length > 1 && m.level > reusable.level && matchesAt(m.numId, reusable.level)));
        if (keepSublevel) continue;
        ops.push({
          op: 'setListNumbering',
          paragraphId,
          numId: reusable.numId,
          level: reusable.level,
        });
      }
      return ops;
    }
  }

  // 3. Mint a fresh list instance (preset switch, new list, or explicit restart) preserving sublevels.
  const numId = deps.ensureListDefinition('ordered', undefined, targetPreset);
  if (!numId) return null;
  const existingLevels = markers
    .filter((m): m is ListMarkerRecord => m !== null)
    .map((m) => m.level);
  const minLevel = existingLevels.length > 0 ? Math.min(...existingLevels) : 0;
  return touched.map((paragraphId, idx) => {
    const m = markers[idx];
    const level =
      m !== null && touched.length > 1
        ? Math.min(MAX_LIST_LEVEL, Math.max(0, m.level - minLevel))
        : 0;
    return {
      op: 'setListNumbering' as const,
      paragraphId,
      numId,
      level,
    };
  });
}

const BORDER_EDGE_ORDER: readonly ParagraphBorderEdgeName[] = [
  'top',
  'left',
  'bottom',
  'right',
  'between',
  'bar',
];
const VALID_BORDER_EDGE_SET: ReadonlySet<string> = new Set<string>(BORDER_EDGE_ORDER);

/**
 * Merge incoming `w:pBdr` edges onto one paragraph's existing direct `w:pBdr`, preserving
 * untouched edges verbatim in strict `CT_PBdr` schema order.
 */
export function mergeParagraphBorderEdges(
  part: OoxmlPart,
  paragraphId: string,
  incoming: readonly ParagraphBorderEdgeWrite[]
): readonly ParagraphBorderEdgeWrite[] {
  const byEdge = new Map<ParagraphBorderEdgeName, Readonly<Record<string, string>>>();
  const paragraph = findNode(part, paragraphId);
  if (paragraph && paragraph.kind !== 'textValue') {
    const pPr = wmlChildNamed(paragraph, 'pPr');
    const pBdr = pPr ? wmlChildNamed(pPr, 'pBdr') : undefined;
    if (pBdr) {
      for (const child of pBdr.children as readonly OoxmlNode[]) {
        if (
          child.kind === 'textValue' ||
          child.namespaceUri !== WML_NS ||
          !VALID_BORDER_EDGE_SET.has(child.localName)
        ) {
          continue;
        }
        const attrs: Record<string, string> = {};
        for (const attr of child.attributes) {
          if (attr.namespaceUri === WML_NS) attrs[attr.localName] = attr.value;
        }
        if (typeof attrs.val === 'string') {
          byEdge.set(child.localName as ParagraphBorderEdgeName, attrs);
        }
      }
    }
  }
  for (const entry of incoming) {
    byEdge.set(entry.edge, { ...entry.attributes });
  }
  const merged: ParagraphBorderEdgeWrite[] = [];
  for (const edge of BORDER_EDGE_ORDER) {
    const attributes = byEdge.get(edge);
    if (attributes) merged.push({ edge, attributes });
  }
  return merged;
}
