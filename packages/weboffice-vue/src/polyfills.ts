/**
 * 国产化 / 信创老内核浏览器运行时 Polyfill（目标向下兼容 Chromium 80～98）。
 * 针对 EigenPal 核心源码扫描出的高版本 JS API：
 * - globalThis.structuredClone (Chrome 98+)
 * - Object.hasOwn (Chrome 93+)
 * - Array.prototype.at / String.prototype.at / TypedArray.prototype.at (Chrome 92+)
 * - String.prototype.replaceAll (Chrome 85+)
 */

export interface PolyfillReport {
  readonly installed: readonly string[];
  readonly alreadyNative: readonly string[];
}

export interface BrowserEnvironmentProbe {
  readonly userAgent: string;
  readonly chromiumMajor: number | null;
  readonly hasWasm: boolean;
  readonly hasWebGL2: boolean;
  readonly hasWebGPU: boolean;
  readonly polyfills: PolyfillReport;
  readonly warnings: readonly string[];
}

function fallbackStructuredClone<T>(value: T, seen = new WeakMap<object, unknown>()): T {
  if (value === null || typeof value !== 'object') {
    return value;
  }
  if (seen.has(value as object)) {
    return seen.get(value as object) as T;
  }
  if (value instanceof Date) {
    return new Date(value.getTime()) as unknown as T;
  }
  if (value instanceof RegExp) {
    return new RegExp(value.source, value.flags) as unknown as T;
  }
  if (value instanceof ArrayBuffer) {
    return value.slice(0) as unknown as T;
  }
  if (ArrayBuffer.isView(value)) {
    const view = value as unknown as ArrayBufferView;
    const raw = new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
    const bufferCopy = new ArrayBuffer(raw.byteLength);
    new Uint8Array(bufferCopy).set(raw);
    const Ctor = (value as unknown as { constructor: new (buf: ArrayBuffer) => T }).constructor;
    return new Ctor(bufferCopy);
  }
  if (value instanceof Map) {
    const out = new Map();
    seen.set(value, out);
    value.forEach((v, k) => {
      out.set(fallbackStructuredClone(k, seen), fallbackStructuredClone(v, seen));
    });
    return out as unknown as T;
  }
  if (value instanceof Set) {
    const out = new Set();
    seen.set(value, out);
    value.forEach((v) => {
      out.add(fallbackStructuredClone(v, seen));
    });
    return out as unknown as T;
  }
  if (Array.isArray(value)) {
    const out: unknown[] = [];
    seen.set(value, out);
    for (let i = 0; i < value.length; i += 1) {
      out[i] = fallbackStructuredClone(value[i], seen);
    }
    return out as unknown as T;
  }
  const out: Record<string, unknown> = {};
  seen.set(value as object, out);
  for (const key of Object.keys(value as Record<string, unknown>)) {
    out[key] = fallbackStructuredClone((value as Record<string, unknown>)[key], seen);
  }
  return out as T;
}

function atMethod<T>(this: ArrayLike<T>, index: number): T | undefined {
  const len = this.length >>> 0;
  let k = Math.trunc(index) || 0;
  if (k < 0) k += len;
  if (k < 0 || k >= len) return undefined;
  return this[k];
}

let cachedReport: PolyfillReport | null = null;

/**
 * 幂等安装老内核浏览器 Polyfill，并返回本次补齐与原生支持清单。
 */
export function installLegacyPolyfills(
  target: Record<string, unknown> = globalThis as Record<string, unknown>
): PolyfillReport {
  if (target === (globalThis as Record<string, unknown>) && cachedReport) {
    return cachedReport;
  }

  const installed: string[] = [];
  const alreadyNative: string[] = [];

  if (typeof target.structuredClone !== 'function') {
    target.structuredClone = <T>(v: T) => fallbackStructuredClone(v);
    installed.push('structuredClone');
  } else {
    alreadyNative.push('structuredClone');
  }

  const objCtor = Object as unknown as Record<string, unknown>;
  if (typeof objCtor.hasOwn !== 'function') {
    Object.defineProperty(Object, 'hasOwn', {
      value(obj: object, prop: PropertyKey) {
        if (obj === null || obj === undefined) {
          throw new TypeError('Cannot convert undefined or null to object');
        }
        return Object.prototype.hasOwnProperty.call(Object(obj), prop);
      },
      configurable: true,
      writable: true,
    });
    installed.push('Object.hasOwn');
  } else {
    alreadyNative.push('Object.hasOwn');
  }

  if (typeof Array.prototype.at !== 'function') {
    Object.defineProperty(Array.prototype, 'at', {
      value: atMethod,
      configurable: true,
      writable: true,
    });
    installed.push('Array.prototype.at');
  } else {
    alreadyNative.push('Array.prototype.at');
  }

  if (typeof String.prototype.at !== 'function') {
    Object.defineProperty(String.prototype, 'at', {
      value: atMethod,
      configurable: true,
      writable: true,
    });
    installed.push('String.prototype.at');
  } else {
    alreadyNative.push('String.prototype.at');
  }

  const typedArrayProto = Object.getPrototypeOf(Uint8Array.prototype) as { at?: unknown };
  if (typedArrayProto && typeof typedArrayProto.at !== 'function') {
    Object.defineProperty(typedArrayProto, 'at', {
      value: atMethod,
      configurable: true,
      writable: true,
    });
    installed.push('TypedArray.prototype.at');
  } else {
    alreadyNative.push('TypedArray.prototype.at');
  }

  if (typeof String.prototype.replaceAll !== 'function') {
    Object.defineProperty(String.prototype, 'replaceAll', {
      value(
        this: string,
        search: string | RegExp,
        replaceValue: string | ((...args: unknown[]) => string)
      ) {
        if (search instanceof RegExp) {
          if (!search.global) {
            throw new TypeError(
              'String.prototype.replaceAll called with a non-global RegExp argument'
            );
          }
          return this.replace(search, replaceValue as string);
        }
        const needle = String(search);
        if (needle === '') {
          const chars = Array.from(this);
          const rep =
            typeof replaceValue === 'function'
              ? String(replaceValue('', 0, this))
              : String(replaceValue);
          return rep + chars.join(rep) + rep;
        }
        return this.split(needle).join(
          typeof replaceValue === 'function'
            ? String(replaceValue(needle, 0, this))
            : String(replaceValue)
        );
      },
      configurable: true,
      writable: true,
    });
    installed.push('String.prototype.replaceAll');
  } else {
    alreadyNative.push('String.prototype.replaceAll');
  }

  const report: PolyfillReport = { installed, alreadyNative };
  if (target === (globalThis as Record<string, unknown>)) {
    cachedReport = report;
  }
  return report;
}

/**
 * G0 / T0 国产化浏览器探针：检测 Chromium 内核、WASM、WebGL2 及 Polyfill 状态。
 */
export function probeBrowserEnvironment(): BrowserEnvironmentProbe {
  const polyfills = installLegacyPolyfills();
  const ua = typeof navigator !== 'undefined' ? navigator.userAgent : 'non-browser';
  const match = ua.match(/(?:Chrome|Chromium|Edg)\/(\d+)/i);
  const chromiumMajor = match ? Number(match[1]) : null;
  const hasWasm = typeof WebAssembly === 'object' && typeof WebAssembly.instantiate === 'function';
  const hasWebGPU = typeof navigator !== 'undefined' && 'gpu' in navigator;

  let hasWebGL2 = false;
  if (typeof document !== 'undefined') {
    try {
      const canvas = document.createElement('canvas');
      hasWebGL2 = Boolean(canvas.getContext('webgl2'));
    } catch {
      hasWebGL2 = false;
    }
  }

  const warnings: string[] = [];
  if (!hasWasm) {
    warnings.push('当前浏览器不支持 WebAssembly，HarfBuzz 排版测量将不可用');
  }
  if (chromiumMajor !== null && chromiumMajor < 86) {
    warnings.push(
      `检测到较老 Chromium 内核 (${chromiumMajor})，请务必启用 legacy 转译包并在真机完成 T0/T6 验收`
    );
  }
  if (polyfills.installed.length > 0) {
    warnings.push(`已启用老内核运行时补丁: ${polyfills.installed.join(', ')}`);
  }

  return {
    userAgent: ua,
    chromiumMajor,
    hasWasm,
    hasWebGL2,
    hasWebGPU,
    polyfills,
    warnings,
  };
}

// 模块导入时即自动补齐运行时缺失 API，保证在任何 core 模块执行前生效
installLegacyPolyfills();
