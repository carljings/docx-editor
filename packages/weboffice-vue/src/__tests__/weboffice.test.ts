import { describe, expect, test } from 'bun:test';
import { installLegacyPolyfills, probeBrowserEnvironment } from '../polyfills';
import { CN_FONT_SIZES, CN_OFFICIAL_FONT_FILES, createOfficialFontResolver } from '../fonts';
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
    expect(bundle.getDiagnostics()).toEqual({
      status: 'idle',
      pendingCount: 0,
      requestedFamilies: [],
      hits: [],
      misses: [],
    });
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
