// Section columns (`w:cols`) validation & surgical element merge for `setSectionProperties`.

import type { OoxmlNode } from '../package/ooxml-tree.ts';

const WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface SectionColumnsWrite {
  readonly count: number;
  readonly spaceTwips?: number;
  readonly separator?: boolean;
}

export const MAX_SECTION_COLUMNS = 12;
export const MAX_COLUMN_SPACE_TWIPS = 7_200;
export const DEFAULT_COLUMN_SPACE_TWIPS = 425;

export function validateSectionColumnsWrite(columns: SectionColumnsWrite): string | null {
  if (
    !Number.isInteger(columns.count) ||
    columns.count < 1 ||
    columns.count > MAX_SECTION_COLUMNS
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
  return null;
}

function wmlAttr(localName: string, value: string) {
  return {
    kind: 'genericExtension' as const,
    namespaceUri: WML_NS,
    localName,
    prefix: 'w',
    value,
  };
}

/** Build a schema-valid `<w:cols>` element matching Word/WPS equal-width column authoring. */
export function buildSectionColsElement(id: string, columns: SectionColumnsWrite): OoxmlNode {
  const space = columns.spaceTwips ?? DEFAULT_COLUMN_SPACE_TWIPS;
  const attrs = [
    ...(columns.count > 1 ? [wmlAttr('num', String(columns.count))] : []),
    wmlAttr('space', String(space)),
    ...(columns.count > 1 && columns.separator ? [wmlAttr('sep', '1')] : []),
  ];
  return {
    id,
    kind: 'generic',
    namespaceUri: WML_NS,
    localName: 'cols',
    prefix: 'w',
    namespaceBindings: [],
    attributes: attrs,
    children: [],
  } as unknown as OoxmlNode;
}
