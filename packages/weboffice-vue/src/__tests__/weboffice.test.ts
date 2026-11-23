import { describe, expect, test } from 'bun:test';
import { installLegacyPolyfills, probeBrowserEnvironment } from '../polyfills';
import {
  CN_FONT_SIZES,
  CN_OFFICIAL_FONT_FILES,
  CN_SELECTABLE_FONT_FAMILIES,
  createOfficialFontResolver,
} from '../fonts';
import {
  CN_OFFICIAL_SYMBOLS,
  GB_T_9704_MARGINS_CM,
  charsToTwips,
  cmToTwips,
  countDocumentWords,
  stepChineseFontSize,
  twipsToChars,
  twipsToCm,
} from '../units';
import { createWebOfficeAdapter } from '../adapter';
import { buildHeaderFooterDistanceUpdate, buildPageSetupUpdate } from '../official-dialog-helpers';
import {
  CHAR_SPACING_PRESETS,
  CN_NUMBERING_PRESETS,
  PARAGRAPH_BORDER_PRESETS,
  PARAGRAPH_SHADING_SWATCHES,
  UNDERLINE_STYLE_PRESETS,
} from '../p1-presets';
import {
  OFFICIAL_WATERMARK_PRESETS,
  TABLE_CELL_FILL_SWATCHES,
  WATERMARK_COLOR_OPTIONS,
  buildCustomWatermark,
  extractOfficialFallbackHeadings,
} from '../p2-presets';

describe('@weboffice/vue polyfills & Xinchuang browser probe', () => {
  test('installs fallback structuredClone on simulated Chromium 86 global target', () => {
    const fakeGlobal: Record<string, unknown> = {};
    const report = installLegacyPolyfills(fakeGlobal);
    expect(report.installed).toContain('structuredClone');

    const clone = fakeGlobal.structuredClone as <T>(v: T) => T;
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const orig = {
      name: '日盯办',
      bytes,
      when: new Date('2026-11-21T00:00:00Z'),
      nested: [{ ok: true }],
    };
    const copied = clone(orig);
    expect(copied).toEqual(orig);
    expect(copied).not.toBe(orig);
    expect(copied.bytes).not.toBe(orig.bytes);
    copied.bytes[0] = 99;
    expect(orig.bytes[0]).toBe(1);
  });

  test('probeBrowserEnvironment returns structured telemetry', () => {
    const env = probeBrowserEnvironment();
    expect(typeof env.hasWasm).toBe('boolean');
    expect(typeof env.hasWebGL2).toBe('boolean');
    expect(Array.isArray(env.warnings)).toBe(true);
  });
});

describe('@weboffice/vue Chinese official document fonts & sizes', () => {
  test('covers full GB/T 9704-2012 Chinese font size range (初号 to 八号)', () => {
    expect(CN_FONT_SIZES[0]).toEqual({ label: '初号', pt: 42 });
    expect(CN_FONT_SIZES.find((s) => s.label === '三号')).toEqual({ label: '三号', pt: 16 });
    expect(CN_FONT_SIZES.find((s) => s.label === '小四')).toEqual({ label: '小四', pt: 12 });
    expect(CN_FONT_SIZES[CN_FONT_SIZES.length - 1]).toEqual({ label: '八号', pt: 5 });
  });

  test('includes required Chinese official font families and creates resolver safely', () => {
    const families = CN_OFFICIAL_FONT_FILES.map((f) => f.family);
    expect(families).toContain('方正小标宋_GBK');
    expect(families).toContain('方正仿宋_GBK');
    expect(families).toContain('仿宋_GB2312');

    const bundle = createOfficialFontResolver({ baseUrl: '/reportcenter/' });
    expect(typeof bundle.resolver).toBe('function');
    expect(CN_SELECTABLE_FONT_FAMILIES).toContain('仿宋_GB2312');
    expect(CN_SELECTABLE_FONT_FAMILIES).toContain('方正小标宋_GBK');
    expect(bundle.getDiagnostics()).toEqual({
      status: 'idle',
      pendingCount: 0,
      requestedFamilies: [],
      hits: [],
      misses: [],
    });
  });

  test('exposes supportedFamilies from official resolver even without network fetches', async () => {
    const bundle = createOfficialFontResolver({
      baseUrl: '/reportcenter/',
      enableFandolFallback: false,
      enablePackagedFonts: false,
    });
    const res = await bundle.resolver({
      families: ['UnknownFont'],
      faces: [],
      defaultFamily: 'Times New Roman',
    });
    expect(res?.supportedFamilies).toContain('仿宋_GB2312');
    expect(res?.supportedFamilies).toContain('黑体');
    expect(res?.supportedFamilies).toContain('方正小标宋_GBK');
  });
});

describe('@weboffice/vue official document units, font ladder, and word statistics', () => {
  test('converts twips <-> cm and chars accurately for GB/T 9704 official formatting', () => {
    // A4 21.0 x 29.7 cm
    expect(twipsToCm(11906)).toBeCloseTo(21.0, 1);
    expect(twipsToCm(16838)).toBeCloseTo(29.7, 1);
    expect(cmToTwips(GB_T_9704_MARGINS_CM.top)).toBe(2098);
    expect(cmToTwips(GB_T_9704_MARGINS_CM.bottom)).toBe(1984);
    expect(cmToTwips(GB_T_9704_MARGINS_CM.left)).toBe(1587);
    expect(cmToTwips(GB_T_9704_MARGINS_CM.right)).toBe(1474);

    // 三号 (16pt) 首行缩进 2 字符 = 32pt = 640 twips
    expect(charsToTwips(2, 16)).toBe(640);
    expect(twipsToChars(640, 16)).toBe(2);
    // 小四 (12pt) 首行缩进 2 字符 = 24pt = 480 twips
    expect(charsToTwips(2, 12)).toBe(480);
  });

  test('steps Chinese font sizes strictly along GB/T 9704 ladder', () => {
    // 三号 (16pt = 32hp) -> 小二 (18pt = 36hp) / 小三 (15pt = 30hp)
    expect(stepChineseFontSize(32, 'increase')).toBe(36);
    expect(stepChineseFontSize(32, 'decrease')).toBe(30);
    // Intermediate non-standard 17pt (34hp) steps to nearest upper/lower standard rung
    expect(stepChineseFontSize(34, 'increase')).toBe(36);
    expect(stepChineseFontSize(34, 'decrease')).toBe(32);
    // Boundaries clamp cleanly
    expect(stepChineseFontSize(84, 'increase')).toBeNull();
    expect(stepChineseFontSize(10, 'decrease')).toBeNull();
    // Null defaults around 三号 (16pt)
    expect(stepChineseFontSize(null, 'increase')).toBe(36);
    expect(stepChineseFontSize(null, 'decrease')).toBe(28);
  });

  test('counts Chinese characters, Latin words, and paragraphs accurately', () => {
    const stats = countDocumentWords([
      { text: '执法监督管理“日盯办”报告' },
      { text: 'ABCDE 2026年第12期' },
      { text: '' },
    ]);
    expect(stats.paragraphCount).toBe(3);
    expect(stats.cjkChars).toBe(14);
    expect(stats.latinWords).toBe(3); // ABCDE, 2026, 12
    expect(stats.totalCharsNoSpace).toBe(27);
    expect(stats.totalCharsWithSpace).toBe(28);
    expect(CN_OFFICIAL_SYMBOLS.some((s) => s.char === '〔')).toBe(true);
    expect(CN_OFFICIAL_SYMBOLS.some((s) => s.char === '〕')).toBe(true);
  });
});

describe('@weboffice/vue engine-agnostic adapter', () => {
  test('reports editor-not-ready cleanly when ref is null', () => {
    const adapter = createWebOfficeAdapter({ getRef: () => null });
    const res = adapter.insertTextAtCursor('{{接报警情总数}}');
    expect(res.ok).toBe(false);
    expect(res.mode).toBe('failed');
    expect(res.reason).toBe('editor-not-ready');
  });

  test('strictly refuses silent fallback by default and supports explicit fallback with position', () => {
    const calls: unknown[] = [];
    let attempt = 0;
    const mockEditor = {
      focus: () => calls.push('focus'),
      getSelectionRange: () => (attempt >= 2 ? { head: { blockId: 'p1', offset: 15 } } : null),
      exec: (cmd: { type: string; text?: string }) => {
        calls.push(cmd);
        if (cmd.type === 'replaceText') {
          attempt += 1;
          return attempt <= 2
            ? { ok: false, changed: false, reason: 'no-selection' }
            : { ok: true, changed: true };
        }
        return { ok: true, changed: false };
      },
      getParagraphs: () => [{ blockId: 'p1', text: '一、警情' }],
    };
    const adapter = createWebOfficeAdapter({
      getRef: () =>
        ({
          getEditor: () => mockEditor,
          load: () => {},
          save: async () => new Uint8Array([80, 75]).buffer,
          focus: () => {},
          exec: () => ({ ok: true, changed: false }),
          snapshot: () => ({}),
          getDocumentHandle: () => null,
        }) as unknown as ReturnType<
          NonNullable<Parameters<typeof createWebOfficeAdapter>[0]['getRef']>
        >,
    });

    // Strict mode (default): no silent fallback
    const strictRes = adapter.insertTextAtCursor('【指标：接报警情总数】');
    expect(strictRes.ok).toBe(false);
    expect(strictRes.mode).toBe('failed');
    expect(strictRes.fallbackApplied).toBe(false);

    // Explicit fallback enabled
    const fallbackRes = adapter.insertTextAtCursor('【指标：接报警情总数】', {
      allowFallback: true,
    });
    expect(fallbackRes.ok).toBe(true);
    expect(fallbackRes.changed).toBe(true);
    expect(fallbackRes.mode).toBe('fallback-first-paragraph');
    expect(fallbackRes.fallbackApplied).toBe(true);
    expect(fallbackRes.position).toEqual({ blockId: 'p1', offset: 15 });
    expect(calls.length).toBeGreaterThanOrEqual(4);
  });
});

describe('@weboffice/vue P1 official presets & page columns dirty-tracking', () => {
  test('exposes GB/T 9704 numbering, official red borders, shading, underline and char spacing presets', () => {
    expect(CN_NUMBERING_PRESETS.map((p) => p.preset)).toEqual([
      'cn-heading1',
      'cn-heading2',
      'cn-decimal',
      'cn-paren',
    ]);
    const red = PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'official-red');
    expect(red?.mode).toBe('merge');
    expect(red?.edges).toEqual([
      { edge: 'bottom', attributes: { val: 'single', sz: '18', space: '4', color: 'DC2626' } },
    ]);
    const doubleRed = PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'official-double-red');
    expect(doubleRed?.mode).toBe('merge');
    expect(doubleRed?.edges[0]?.attributes.val).toBe('double');
    expect(PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'bottom-thin')?.mode).toBe('merge');
    expect(PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'box')?.edges.length).toBe(4);
    expect(PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'box')?.mode).toBe('replace');
    expect(PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'none')?.edges.length).toBe(0);
    expect(PARAGRAPH_BORDER_PRESETS.find((b) => b.id === 'none')?.mode).toBe('replace');
    for (const s of PARAGRAPH_SHADING_SWATCHES) {
      expect(s.hex === null || /^[0-9A-F]{6}$/.test(s.hex)).toBe(true);
    }
    expect(UNDERLINE_STYLE_PRESETS.some((u) => u.value === 'wave')).toBe(true);
    expect(UNDERLINE_STYLE_PRESETS.some((u) => u.value === 'none')).toBe(true);
    for (const cs of CHAR_SPACING_PRESETS) {
      expect(cs.twips).toBe(Math.round(cs.pt * 20));
    }
  });

  test('buildPageSetupUpdate patches columns per-attribute and supports zero-gap and 4+ columns', () => {
    const base = {
      dirty: {
        paper: false,
        orientation: false,
        marginTop: false,
        marginBottom: false,
        marginLeft: false,
        marginRight: false,
        columnsCount: false,
        columnsSpace: false,
        columnsSep: false,
      },
      paperId: 'custom',
      customPaperSize: { widthTwips: 12500, heightTwips: 17000 },
      orientation: 'portrait' as const,
      marginTopStr: '3.7',
      marginBottomStr: '3.5',
      marginLeftStr: '2.8',
      marginRightStr: '2.6',
      columnsCount: 2,
      columnsSpaceStr: '0.75',
      columnsSeparator: true,
      scope: 'section' as const,
    };

    // Only marginTop dirty -> columns & paper dimensions omitted
    const marginOnly = buildPageSetupUpdate({
      ...base,
      dirty: { ...base.dirty, marginTop: true },
    });
    expect(marginOnly.hasChanges).toBe(true);
    expect(marginOnly.update.columns).toBeUndefined();
    expect(marginOnly.update.pageWidthTwips).toBeUndefined();

    // Only separator dirty -> emits ONLY { separator: true } without clobbering count/space/equalWidth
    const sepOnly = buildPageSetupUpdate({
      ...base,
      columnsCount: 4,
      columnsSpaceStr: '0',
      dirty: { ...base.dirty, columnsSep: true },
    });
    expect(sepOnly.hasChanges).toBe(true);
    expect(sepOnly.update.columns).toEqual({ separator: true });

    // Explicit zero gap -> emits spaceTwips: 0
    const zeroGap = buildPageSetupUpdate({
      ...base,
      columnsSpaceStr: '0',
      dirty: { ...base.dirty, columnsSpace: true },
    });
    expect(zeroGap.hasChanges).toBe(true);
    expect(zeroGap.update.columns).toEqual({ spaceTwips: 0 });

    // All three column fields dirty -> emits full columns payload without overwriting paper/margins
    const colsAll = buildPageSetupUpdate({
      ...base,
      dirty: { ...base.dirty, columnsCount: true, columnsSpace: true, columnsSep: true },
    });
    expect(colsAll.hasChanges).toBe(true);
    expect(colsAll.update.pageWidthTwips).toBeUndefined();
    expect(colsAll.update.marginTopTwips).toBeUndefined();
    expect(colsAll.update.columns).toEqual({
      count: 2,
      spaceTwips: cmToTwips(0.75),
      separator: true,
    });
    expect(colsAll.update.scope).toBe('section');

    // Invalid column gap rejected
    const badGap = buildPageSetupUpdate({
      ...base,
      dirty: { ...base.dirty, columnsSpace: true },
      columnsSpaceStr: '12',
    });
    expect(badGap.error).toBeDefined();
    expect(badGap.hasChanges).toBe(false);
  });
});

describe('@weboffice/vue P2 presets, fallback outline & header/footer distances', () => {
  test('provides official Chinese watermark presets and table cell fill swatches', () => {
    expect(OFFICIAL_WATERMARK_PRESETS.map((p) => p.label)).toEqual([
      '内部资料',
      '机密',
      '草稿 · 仅供核稿',
      '严禁外传',
    ]);
    expect(WATERMARK_COLOR_OPTIONS.length).toBeGreaterThanOrEqual(4);
    expect(TABLE_CELL_FILL_SWATCHES[0]?.hex).toBeNull();
    expect(TABLE_CELL_FILL_SWATCHES.length).toBe(7);
  });

  test('buildCustomWatermark validates, normalizes whitespace, rejects over-length, and falls back color', () => {
    const empty = buildCustomWatermark({ text: '   ', layout: 'diagonal', color: 'C0C0C0' });
    expect(empty.ok).toBe(false);

    const custom = buildCustomWatermark({
      text: '  执法监督\n  内部核稿  ',
      layout: 'horizontal',
      color: 'dc2626',
    });
    expect(custom.ok).toBe(true);
    if (custom.ok) {
      expect(custom.watermark).toMatchObject({
        text: '执法监督 内部核稿',
        fontFamily: '宋体',
        layout: 'horizontal',
        color: 'DC2626',
        transparency: 0.5,
      });
    }

    const tooLong = buildCustomWatermark({
      text: '甲'.repeat(80),
      layout: 'diagonal',
      color: 'C0C0C0',
    });
    expect(tooLong.ok).toBe(false);

    const fallbackHex = buildCustomWatermark({
      text: '内部参考',
      layout: 'diagonal',
      color: 'not-hex',
    });
    expect(fallbackHex.ok).toBe(true);
    if (fallbackHex.ok) {
      expect(fallbackHex.watermark.color).toBe('C0C0C0');
    }
  });

  test('extractOfficialFallbackHeadings identifies 一、 and （一） official headings while skipping body paragraphs', () => {
    const items = extractOfficialFallbackHeadings([
      { blockId: 'b1', text: '   ' },
      { blockId: 'b2', text: '一、总体工作情况' },
      { blockId: 'b3', text: '（一）警情受理与处置效率提升' },
      {
        blockId: 'b4',
        text: '一、这是一段超过六十四个字符的正文叙述，明显不属于公文小节标题，应该被大纲自动识别逻辑过滤掉，避免把大段公文正文误判成左侧导航目录里的标题条目。',
      },
      { blockId: 'b5', text: '二、下一步工作措施' },
      { blockId: 'b6', text: '(二) 强化闭环督办' },
    ]);
    expect(items).toEqual([
      { blockId: 'b2', depth: 0, text: '一、总体工作情况' },
      { blockId: 'b3', depth: 1, text: '（一）警情受理与处置效率提升' },
      { blockId: 'b5', depth: 0, text: '二、下一步工作措施' },
      { blockId: 'b6', depth: 1, text: '(二) 强化闭环督办' },
    ]);
  });

  test('buildHeaderFooterDistanceUpdate tracks independent dirty flags and validates 0..15 cm range', () => {
    const clean = buildHeaderFooterDistanceUpdate({
      headerDirty: false,
      footerDirty: false,
      headerDistStr: '1.5',
      footerDistStr: '1.75',
    });
    expect(clean.hasChanges).toBe(false);

    const headerOnly = buildHeaderFooterDistanceUpdate({
      headerDirty: true,
      footerDirty: false,
      headerDistStr: '1.5',
      footerDistStr: '99',
    });
    expect(headerOnly.hasChanges).toBe(true);
    expect(headerOnly.update.headerDistanceTwips).toBe(cmToTwips(1.5));
    expect(headerOnly.update.footerDistanceTwips).toBeUndefined();

    const both = buildHeaderFooterDistanceUpdate({
      headerDirty: true,
      footerDirty: true,
      headerDistStr: '1.2',
      footerDistStr: '1.8',
    });
    expect(both.hasChanges).toBe(true);
    expect(both.update.headerDistanceTwips).toBe(cmToTwips(1.2));
    expect(both.update.footerDistanceTwips).toBe(cmToTwips(1.8));

    const outOfRange = buildHeaderFooterDistanceUpdate({
      headerDirty: false,
      footerDirty: true,
      headerDistStr: '1.5',
      footerDistStr: '16',
    });
    expect(outOfRange.hasChanges).toBe(false);
    expect(outOfRange.error).toBeDefined();
  });
});
