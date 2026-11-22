import { packagedFonts } from '@docx-editor.dev/fonts';
import {
  composeFontOrigins,
  customFonts,
  defineFontResolver,
  type FontOrigin,
  type FontUrlSource,
} from '@docx-editor.dev/core/editor';

export interface ChineseFontSizeOption {
  readonly label: string;
  readonly pt: number;
}

/**
 * 中国党政机关公文及 WPS / MS Word 常用字号表（初号 42pt ～ 八号 5pt）。
 */
export const CN_FONT_SIZES: readonly ChineseFontSizeOption[] = [
  { label: '初号', pt: 42 },
  { label: '小初', pt: 36 },
  { label: '一号', pt: 26 },
  { label: '小一', pt: 24 },
  { label: '二号', pt: 22 },
  { label: '小二', pt: 18 },
  { label: '三号', pt: 16 },
  { label: '小三', pt: 15 },
  { label: '四号', pt: 14 },
  { label: '小四', pt: 12 },
  { label: '五号', pt: 10.5 },
  { label: '小五', pt: 9 },
  { label: '六号', pt: 7.5 },
  { label: '小六', pt: 6.5 },
  { label: '七号', pt: 5.5 },
  { label: '八号', pt: 5 },
];

export interface OfficialFontFileSpec {
  readonly family: string;
  readonly file: string;
  readonly weight?: number;
  readonly style?: 'normal' | 'italic';
}

export interface FontStandInSpec {
  readonly family: string;
  readonly file: string;
  readonly originals: readonly string[];
}

export interface FontHitRecord {
  readonly layer: 'official' | 'wps-local' | 'fandol' | 'packaged';
  readonly family: string;
  readonly url?: string;
  readonly substitutedFor?: readonly string[];
}

export interface FontMissRecord {
  readonly layer: 'official' | 'wps-local' | 'fandol';
  readonly family: string;
  readonly url: string;
  readonly reason: string;
}

export type FontDiagnosticsStatus = 'idle' | 'loading' | 'settled';

export interface FontDiagnosticsSnapshot {
  readonly status: FontDiagnosticsStatus;
  readonly pendingCount: number;
  readonly requestedFamilies: readonly string[];
  readonly hits: readonly FontHitRecord[];
  readonly misses: readonly FontMissRecord[];
}

export const CN_OFFICIAL_FONT_FILES: readonly OfficialFontFileSpec[] = [
  { family: '仿宋', file: '仿宋.ttf' },
  { family: '仿宋_GB2312', file: '仿宋_GB2312.ttf' },
  { family: '方正仿宋_GBK', file: '方正仿宋_GBK.ttf' },
  { family: '方正小标宋_GBK', file: '方正小标宋_GBK.ttf' },
  { family: '方正楷体_GBK', file: '方正楷体_GBK.ttf' },
  { family: '楷体_GB2312', file: '楷体_GB2312.ttf' },
  { family: '楷体', file: '楷体.ttf' },
  { family: '方正黑体_GBK', file: '方正黑体_GBK.ttf' },
  { family: '黑体', file: '黑体.ttf' },
  { family: '宋体', file: '宋体.ttf' },
];

/**
 * 字体下拉默认展示的中文公文 + 常用西文字体目录（声明后不触发预下载，仅在选用时按需解析）。
 */
export const CN_SELECTABLE_FONT_FAMILIES: readonly string[] = Object.freeze([
  '仿宋_GB2312',
  '仿宋',
  '方正仿宋_GBK',
  '方正小标宋_GBK',
  '黑体',
  '方正黑体_GBK',
  '宋体',
  '楷体_GB2312',
  '楷体',
  '方正楷体_GBK',
  'Times New Roman',
  'Calibri',
  'Arial',
]);

/**
 * 开源 Fandol 0.3 兜底字族（无商业授权风险，适合内网未购方正字库时的合规兜底）。
 */
export const FANDOL_STAND_INS: readonly FontStandInSpec[] = [
  {
    family: 'FandolFang',
    file: 'FandolFang-Regular.otf',
    originals: ['仿宋', '仿宋_GB2312', '方正仿宋_GBK', '仿宋-简'],
  },
  {
    family: 'FandolSong',
    file: 'FandolSong-Bold.otf',
    originals: ['方正小标宋_GBK', '宋体', '书宋-简', '宋体-简'],
  },
  {
    family: 'FandolKai',
    file: 'FandolKai-Regular.otf',
    originals: ['方正楷体_GBK', '楷体_GB2312', '楷体', '汉仪楷体简'],
  },
  {
    family: 'FandolHei',
    file: 'FandolHei-Regular.otf',
    originals: ['方正黑体_GBK', '黑体', '汉仪中简黑简'],
  },
];

/**
 * 仅限本机开发对齐 WPS 用的替代映射；无网页分发权，生产环境默认禁用。
 */
export const WPS_LOCAL_DEV_STAND_INS: readonly FontStandInSpec[] = [
  {
    family: 'FangS-SC',
    file: 'FangS-SC.ttf',
    originals: ['仿宋', '仿宋_GB2312', '方正仿宋_GBK', '仿宋-简'],
  },
  { family: 'ShuS-SC', file: 'ShuS-SC.ttf', originals: ['宋体', '书宋-简', '宋体-简'] },
  {
    family: 'HYKaiTiJ',
    file: 'HYKaiTiJ.ttf',
    originals: ['方正楷体_GBK', '楷体_GB2312', '楷体', '汉仪楷体简'],
  },
];

const LATIN_VARIANTS = [
  { family: 'Calibri', file: 'calibri.ttf', weight: 400, style: 'normal' as const },
  { family: 'Calibri', file: 'calibrib.ttf', weight: 700, style: 'normal' as const },
  { family: 'Calibri', file: 'calibrii.ttf', weight: 400, style: 'italic' as const },
  { family: 'Calibri', file: 'calibriz.ttf', weight: 700, style: 'italic' as const },
  { family: 'Times New Roman', file: 'times.ttf', weight: 400, style: 'normal' as const },
  { family: 'Times New Roman', file: 'timesbd.ttf', weight: 700, style: 'normal' as const },
  { family: 'Times New Roman', file: 'timesi.ttf', weight: 400, style: 'italic' as const },
  { family: 'Times New Roman', file: 'timesbi.ttf', weight: 700, style: 'italic' as const },
];

export interface OfficialFontResolverOptions {
  /** 静态资源基路径，如 `'/'` 或 `import.meta.env.BASE_URL` */
  readonly baseUrl?: string;
  /** 字体目录名，默认 `'fonts'` */
  readonly fontsDir?: string;
  /** 客户授权的原版公文字体列表，默认 `CN_OFFICIAL_FONT_FILES` */
  readonly officialFiles?: readonly OfficialFontFileSpec[];
  /** 是否自动尝试加载官方目录的 Calibri / Times New Roman，默认 `true` */
  readonly includeOfficialLatin?: boolean;
  /** 是否启用开源 Fandol 兜底字体（位于 `<fontsDir>/fandol/`），默认 `true` */
  readonly enableFandolFallback?: boolean;
  /**
   * 仅限本机开发对齐测试：是否启用 `<fontsDir>/wps/` 下的本机 WPS 字体。
   * 生产环境必须保持默认 `false`（WPS 桌面字形无网页部署授权）。
   */
  readonly enableWpsLocalStandIns?: boolean;
  /** 是否启用 EigenPal 内置离线开源西文字体包，默认 `true` */
  readonly enablePackagedFonts?: boolean;
  /** 字体解析命中/缺失诊断回调，便于 G0/G1 验收追踪 */
  readonly onDiagnostics?: (snapshot: FontDiagnosticsSnapshot) => void;
}

function normalizeBaseUrl(baseUrl: string): string {
  return baseUrl.endsWith('/') ? baseUrl : `${baseUrl}/`;
}

function unresolvedFamilies(request: {
  readonly families: readonly string[];
  readonly resolvedFaces?: readonly { readonly family: string }[];
}): Set<string> {
  const resolved = new Set(
    (request.resolvedFaces ?? []).map((face) => face.family.trim().toLowerCase())
  );
  const missing = new Set<string>();
  for (const family of request.families) {
    const norm = family.trim().toLowerCase();
    if (norm && !resolved.has(norm)) missing.add(norm);
  }
  return missing;
}

/**
 * 构建纯内网离线的公文多层字体解析器：
 * 1. 客户授权原版公文字体（`<baseUrl><fontsDir>/`）
 * 2. （仅本机开发显式开启时）WPS 替代字体（`<baseUrl><fontsDir>/wps/`）
 * 3. 开源 Fandol 0.3 兜底字体（`<baseUrl><fontsDir>/fandol/`）
 * 4. EigenPal 内置离线西文字体包（`@docx-editor.dev/fonts`）
 */
export function createOfficialFontResolver(options: OfficialFontResolverOptions = {}) {
  const baseUrl = normalizeBaseUrl(options.baseUrl ?? '/');
  const fontsDir = (options.fontsDir ?? 'fonts').replace(/^\/+|\/+$/g, '');
  const rootPrefix = `${baseUrl}${fontsDir}/`;
  const officialFiles = options.officialFiles ?? CN_OFFICIAL_FONT_FILES;
  const includeOfficialLatin = options.includeOfficialLatin ?? true;
  const enableFandolFallback = options.enableFandolFallback ?? true;
  const enableWpsLocalStandIns = options.enableWpsLocalStandIns ?? false;
  const enablePackagedFonts = options.enablePackagedFonts ?? true;

  const hits: FontHitRecord[] = [];
  const misses: FontMissRecord[] = [];
  const hitKeys = new Set<string>();
  const missKeys = new Set<string>();
  const requestedSet = new Set<string>();
  let pendingCount = 0;
  let totalRequests = 0;

  const getStatus = (): FontDiagnosticsStatus => {
    if (pendingCount > 0) return 'loading';
    return totalRequests > 0 ? 'settled' : 'idle';
  };

  const getSnapshot = (): FontDiagnosticsSnapshot => ({
    status: getStatus(),
    pendingCount,
    requestedFamilies: [...requestedSet],
    hits: [...hits],
    misses: [...misses],
  });

  const emitDiagnostics = () => {
    options.onDiagnostics?.(getSnapshot());
  };

  const recordHit = (record: FontHitRecord) => {
    const key = `${record.layer}:${record.family}:${record.url ?? ''}`;
    if (hitKeys.has(key)) return;
    hitKeys.add(key);
    hits.push(record);
    emitDiagnostics();
  };

  const recordMiss = (record: FontMissRecord) => {
    const key = `${record.layer}:${record.family}:${record.url}`;
    if (missKeys.has(key)) return;
    missKeys.add(key);
    misses.push(record);
    emitDiagnostics();
  };

  const officialSources: FontUrlSource[] = [
    ...officialFiles.map((item) => ({
      family: item.family,
      url: `${rootPrefix}${encodeURIComponent(item.file)}`,
      weight: item.weight ?? 400,
      style: item.style ?? ('normal' as const),
    })),
    ...(includeOfficialLatin
      ? LATIN_VARIANTS.map((face) => ({
          family: face.family,
          url: `${rootPrefix}${face.file}`,
          weight: face.weight,
          style: face.style,
        }))
      : []),
  ];

  const officialOrigin = defineFontResolver(async (request) => {
    const missing = unresolvedFamilies(request);
    if (!missing.size) return undefined;
    const needed = officialSources.filter((src) => missing.has(src.family.trim().toLowerCase()));
    if (!needed.length) return undefined;

    const result = await customFonts({
      sources: needed,
      onFailure: (failure) => {
        recordMiss({
          layer: 'official',
          family: failure.request.family,
          url: failure.url,
          reason: failure.reason,
        });
      },
    })(request);

    for (const src of result?.sources ?? []) {
      const matched = needed.find(
        (item) => item.family.trim().toLowerCase() === src.request.family.trim().toLowerCase()
      );
      recordHit({ layer: 'official', family: src.request.family, url: matched?.url });
    }
    return result?.sources?.length ? result : undefined;
  });

  const createStandInOrigin = (
    layer: 'wps-local' | 'fandol',
    subDir: string,
    standIns: readonly FontStandInSpec[]
  ) =>
    defineFontResolver(async (request) => {
      const missing = unresolvedFamilies(request);
      if (!missing.size) return undefined;
      const requested = standIns
        .map((font) => ({
          ...font,
          neededOriginals: font.originals.filter((family) =>
            missing.has(family.trim().toLowerCase())
          ),
        }))
        .filter((font) => font.neededOriginals.length > 0);
      if (!requested.length) return undefined;

      const result = await customFonts({
        sources: requested.map(({ family, file }) => ({
          family,
          url: `${rootPrefix}${subDir}/${encodeURIComponent(file)}`,
          weight: 400,
          style: 'normal' as const,
        })),
        onFailure: (failure) => {
          recordMiss({
            layer,
            family: failure.request.family,
            url: failure.url,
            reason: failure.reason,
          });
        },
      })(request);

      const loaded = new Set((result?.sources ?? []).map((src) => src.request.family));
      if (!loaded.size) return undefined;

      for (const item of requested) {
        if (loaded.has(item.family)) {
          recordHit({
            layer,
            family: item.family,
            url: `${rootPrefix}${subDir}/${encodeURIComponent(item.file)}`,
            substitutedFor: item.neededOriginals,
          });
        }
      }

      return {
        ...result,
        substitutions: requested
          .filter(({ family }) => loaded.has(family))
          .flatMap(({ family, neededOriginals }) =>
            neededOriginals.flatMap((original) =>
              [400, 700].flatMap((weight) =>
                (['normal', 'italic'] as const).map((style) => ({
                  from: { family: original, weight, style },
                  to: { family, weight: 400, style: 'normal' as const },
                }))
              )
            )
          ),
      };
    });

  const catalogOrigin: FontOrigin = {
    supportedFamilies: CN_SELECTABLE_FONT_FAMILIES,
    sources: [],
  };

  const origins: FontOrigin[] = [catalogOrigin, officialOrigin];
  if (enableWpsLocalStandIns) {
    origins.push(createStandInOrigin('wps-local', 'wps', WPS_LOCAL_DEV_STAND_INS));
  }
  if (enableFandolFallback) {
    origins.push(createStandInOrigin('fandol', 'fandol', FANDOL_STAND_INS));
  }
  if (enablePackagedFonts) {
    const innerPackaged = packagedFonts();
    origins.push(
      defineFontResolver(async (request) => {
        const result = await innerPackaged(request);
        for (const src of result?.sources ?? []) {
          recordHit({ layer: 'packaged', family: src.request.family });
        }
        return result;
      })
    );
  }

  return {
    resolver: defineFontResolver(async (request) => {
      totalRequests += 1;
      pendingCount += 1;
      for (const family of request.families) {
        const trimmed = family.trim();
        if (trimmed) requestedSet.add(trimmed);
      }
      emitDiagnostics();
      try {
        return await composeFontOrigins(origins, request);
      } finally {
        pendingCount = Math.max(0, pendingCount - 1);
        emitDiagnostics();
      }
    }),
    getDiagnostics: getSnapshot,
  };
}
