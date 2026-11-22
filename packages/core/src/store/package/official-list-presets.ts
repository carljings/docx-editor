// Chinese official-document (GB/T 9704-2012) multilevel numbering presets.
// Emits strict `CT_Lvl` schema sequence: start, numFmt, suff, lvlText, lvlJc, pPr.

/** Named ordered-list presets for Chinese official documents (GB/T 9704-2012). */
export type OrderedListPreset = 'cn-heading1' | 'cn-heading2' | 'cn-decimal' | 'cn-paren';

export const ORDERED_LIST_PRESETS: ReadonlySet<string> = new Set<OrderedListPreset>([
  'cn-heading1',
  'cn-heading2',
  'cn-decimal',
  'cn-paren',
]);

interface PresetLevelSpec {
  readonly numFmt: 'chineseCountingThousand' | 'decimal' | 'lowerLetter';
  readonly lvlText: (level1Based: number) => string;
  readonly suffix: 'nothing' | 'space' | 'tab';
}

// GB/T 9704-2012 4-level official hierarchy: 一、 -> （一） -> 1. -> （1）
const CN_HIERARCHY: readonly PresetLevelSpec[] = [
  { numFmt: 'chineseCountingThousand', lvlText: (n) => `%${n}、`, suffix: 'nothing' },
  { numFmt: 'chineseCountingThousand', lvlText: (n) => `（%${n}）`, suffix: 'nothing' },
  { numFmt: 'decimal', lvlText: (n) => `%${n}.`, suffix: 'space' },
  { numFmt: 'decimal', lvlText: (n) => `（%${n}）`, suffix: 'nothing' },
];

const PRESET_START_OFFSET: Record<OrderedListPreset, number> = {
  'cn-heading1': 0,
  'cn-heading2': 1,
  'cn-decimal': 2,
  'cn-paren': 3,
};

/**
 * Build one `w:lvl` element for a Chinese official numbering preset in strict `CT_Lvl` order.
 * Uses 2-character (640 twips at 16pt 三号) first-line indent on level 0 matching GB/T 9704.
 */
export function officialPresetLevelXml(preset: OrderedListPreset, ilvl: number): string {
  const offset = PRESET_START_OFFSET[preset];
  const spec = CN_HIERARCHY[(offset + ilvl) % CN_HIERARCHY.length]!;
  const left = ilvl * 420;
  const firstLine = ilvl === 0 ? 640 : 0;
  const indAttrs =
    firstLine > 0
      ? ` w:left="0" w:firstLine="${firstLine}" w:firstLineChars="200"`
      : ` w:left="${left}" w:hanging="420"`;
  const suffXml = spec.suffix === 'tab' ? '' : `<w:suff w:val="${spec.suffix}"/>`;
  return (
    `<w:lvl w:ilvl="${ilvl}">` +
    `<w:start w:val="1"/>` +
    `<w:numFmt w:val="${spec.numFmt}"/>` +
    suffXml +
    `<w:lvlText w:val="${spec.lvlText(ilvl + 1)}"/>` +
    `<w:lvlJc w:val="left"/>` +
    `<w:pPr><w:ind${indAttrs}/></w:pPr>` +
    '</w:lvl>'
  );
}
