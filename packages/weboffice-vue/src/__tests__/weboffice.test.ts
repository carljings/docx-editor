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
