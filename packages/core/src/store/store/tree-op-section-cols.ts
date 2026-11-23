// Section columns (`w:cols`) validation & lossless attribute/child merge for `setSectionProperties`.

import type { OoxmlAttribute, OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';

const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface SectionColumnsWrite {
  readonly count?: number;
  readonly spaceTwips?: number;
  readonly separator?: boolean;
}

export const MAX_SECTION_COLUMNS = 12;
export const MAX_COLUMN_SPACE_TWIPS = 7_200;
export const DEFAULT_COLUMN_SPACE_TWIPS = 425;

export function validateSectionColumnsWrite(columns: SectionColumnsWrite): string | null {
  if (
    columns.count === undefined &&
    columns.spaceTwips === undefined &&
    columns.separator === undefined
  ) {
    return 'columns requires at least one of count, spaceTwips, or separator';
  }
  if (
    columns.count !== undefined &&
    (!Number.isInteger(columns.count) || columns.count < 1 || columns.count > MAX_SECTION_COLUMNS)
  ) {
    return `columns.count must be an integer between 1 and ${MAX_SECTION_COLUMNS}`;
  }
  if (
    columns.spaceTwips !== undefined &&
    (!Number.isInteger(columns.spaceTwips) ||
      columns.spaceTwips < 0 ||
      columns.spaceTwips > MAX_COLUMN_SPACE_TWIPS)
  ) {
    return `columns.spaceTwips must be an integer between 0 and ${MAX_COLUMN_SPACE_TWIPS}`;
  }
  if (columns.separator !== undefined && typeof columns.separator !== 'boolean') {
    return 'columns.separator must be a boolean';
  }
  return null;
}

function wmlAttr(localName: string, value: string): OoxmlAttribute {
  return {
    kind: 'genericExtension',
    namespaceUri: WML_NS,
    localName,
    prefix: 'w',
    value,
  };
}

function getWmlAttr(node: OoxmlElement | undefined, localName: string): string | undefined {
  if (!node) return undefined;
  return node.attributes.find(
    (attr) => attr.namespaceUri === WML_NS && attr.localName === localName
  )?.value;
}

function isWmlBoolFalse(raw: string | undefined): boolean {
  if (raw === undefined) return false;
  const low = raw.toLowerCase();
  return low === '0' || low === 'false' || low === 'off';
}

function isWmlBoolTrue(raw: string | undefined): boolean {
  if (raw === undefined) return false;
  return !isWmlBoolFalse(raw);
}

function replaceOrRemoveWmlAttr(
  attrs: readonly OoxmlAttribute[],
  localName: string,
  value: string | null
): OoxmlAttribute[] {
  const kept = attrs.filter((a) => !(a.namespaceUri === WML_NS && a.localName === localName));
  if (value !== null) kept.push(wmlAttr(localName, value));
  return kept;
}

function isWmlColNode(node: OoxmlNode): node is OoxmlElement {
  return node.kind !== 'textValue' && node.namespaceUri === WML_NS && node.localName === 'col';
}

function updateExplicitColGaps(children: readonly OoxmlNode[], spaceTwips: number): OoxmlNode[] {
  const colIndices: number[] = [];
  children.forEach((child, idx) => {
    if (isWmlColNode(child)) colIndices.push(idx);
  });
  if (colIndices.length <= 1) return [...children];
  const lastColIdx = colIndices[colIndices.length - 1]!;
  return children.map((child, idx) => {
    if (!isWmlColNode(child) || idx === lastColIdx) return child;
    return {
      ...child,
      attributes: replaceOrRemoveWmlAttr(child.attributes, 'space', String(spaceTwips)),
    } as OoxmlNode;
  });
}

/**
 * Losslessly patch an existing `<w:cols>` (or author a new one).
 *
 * Modifying only `separator` preserves `w:equalWidth="0"`, `w:space="0"`, `w:num`, and every
 * explicit `<w:col w:w="..." w:space="..."/>` child verbatim. Switching `count` to a different
 * column count replaces stale `<w:col>` definitions with equal-width columns.
 */
export function mergeSectionColsElement(
  existingNode: OoxmlNode | null | undefined,
  id: string,
  patch: SectionColumnsWrite
): OoxmlNode {
  const existing =
    existingNode && existingNode.kind !== 'textValue' ? (existingNode as OoxmlElement) : undefined;
  const existingAttrs = existing?.attributes ?? [];
  const existingChildren = existing?.children ?? [];
  const explicitCols = existingChildren.filter(isWmlColNode);
  const rawNum = getWmlAttr(existing, 'num');
  const parsedNum = rawNum && /^\d+$/.test(rawNum) ? Number(rawNum) : undefined;
  const existingCount = parsedNum ?? (explicitCols.length > 0 ? explicitCols.length : 1);
  const existingEqualWidth = !isWmlBoolFalse(getWmlAttr(existing, 'equalWidth'));
  const existingSpace = getWmlAttr(existing, 'space');
  const existingSep = isWmlBoolTrue(getWmlAttr(existing, 'sep'));

  let attrs = [...existingAttrs];
  let children: readonly OoxmlNode[] = existingChildren;

  // 1. Attribute-only patch (count omitted): preserve all existing column geometry & <w:col> children.
  if (patch.count === undefined) {
    if (patch.spaceTwips !== undefined) {
      attrs = replaceOrRemoveWmlAttr(attrs, 'space', String(patch.spaceTwips));
      if (!existingEqualWidth && explicitCols.length > 0) {
        children = updateExplicitColGaps(children, patch.spaceTwips);
      }
    }
    if (patch.separator !== undefined) {
      attrs = replaceOrRemoveWmlAttr(attrs, 'sep', patch.separator ? '1' : null);
    }
    return {
      id,
      kind: 'generic',
      namespaceUri: WML_NS,
      localName: 'cols',
      prefix: existing?.prefix ?? 'w',
      namespaceBindings: existing?.namespaceBindings ?? [],
      attributes: attrs,
      children,
    } as unknown as OoxmlNode;
  }

  // 2. Explicit 1-column reset: clear multi-column attributes and <w:col> children.
  if (patch.count === 1) {
    attrs = replaceOrRemoveWmlAttr(attrs, 'num', null);
    attrs = replaceOrRemoveWmlAttr(attrs, 'sep', null);
    attrs = replaceOrRemoveWmlAttr(attrs, 'equalWidth', null);
    attrs = replaceOrRemoveWmlAttr(
      attrs,
      'space',
      String(patch.spaceTwips ?? existingSpace ?? DEFAULT_COLUMN_SPACE_TWIPS)
    );
    children = existingChildren.filter((child) => !isWmlColNode(child));
    return {
      id,
      kind: 'generic',
      namespaceUri: WML_NS,
      localName: 'cols',
      prefix: existing?.prefix ?? 'w',
      namespaceBindings: existing?.namespaceBindings ?? [],
      attributes: attrs,
      children,
    } as unknown as OoxmlNode;
  }

  // 3. Explicit multi-column count (> 1).
  const keepCustomWidths =
    patch.count === existingCount && !existingEqualWidth && explicitCols.length > 0;
  if (keepCustomWidths) {
    attrs = replaceOrRemoveWmlAttr(attrs, 'num', String(patch.count));
    if (patch.spaceTwips !== undefined) {
      attrs = replaceOrRemoveWmlAttr(attrs, 'space', String(patch.spaceTwips));
      children = updateExplicitColGaps(children, patch.spaceTwips);
    }
    if (patch.separator !== undefined) {
      attrs = replaceOrRemoveWmlAttr(attrs, 'sep', patch.separator ? '1' : null);
    }
  } else {
    const nextSpace =
      patch.spaceTwips !== undefined
        ? String(patch.spaceTwips)
        : existingSpace !== undefined
          ? existingSpace
          : String(DEFAULT_COLUMN_SPACE_TWIPS);
    const nextSep = patch.separator !== undefined ? patch.separator : existingSep;
    attrs = replaceOrRemoveWmlAttr(attrs, 'num', String(patch.count));
    attrs = replaceOrRemoveWmlAttr(attrs, 'space', nextSpace);
    attrs = replaceOrRemoveWmlAttr(attrs, 'equalWidth', null);
    attrs = replaceOrRemoveWmlAttr(attrs, 'sep', nextSep ? '1' : null);
    children = existingChildren.filter((child) => !isWmlColNode(child));
  }

  return {
    id,
    kind: 'generic',
    namespaceUri: WML_NS,
    localName: 'cols',
    prefix: existing?.prefix ?? 'w',
    namespaceBindings: existing?.namespaceBindings ?? [],
    attributes: attrs,
    children,
  } as unknown as OoxmlNode;
}
