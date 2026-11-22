#!/usr/bin/env node
/**
 * Apache-2.0 合规门禁检查：
 * 1. 确保 packages/weboffice-vue 与 examples/weboffice-workbench 零引用四个 EigenPal Pro 专有许可包：
 *    - @docx-editor.dev/pro
 *    - @docx-editor.dev/editor-api
 *    - @docx-editor.dev/docx-to-pdf
 *    - @docx-editor.dev/telemetry
 * 2. 确保未将本机 WPS 桌面字体目录（public/fonts/wps）纳入交付静态资源。
 */

import { readdirSync, readFileSync, statSync, existsSync } from 'node:fs';
import { join, resolve } from 'node:path';

const ROOT = resolve(import.meta.dirname, '..');
const FORBIDDEN_PACKAGES = [
  '@docx-editor.dev/pro',
  '@docx-editor.dev/editor-api',
  '@docx-editor.dev/docx-to-pdf',
  '@docx-editor.dev/telemetry',
];

const SCAN_ROOTS = [
  join(ROOT, 'packages/weboffice-vue'),
  join(ROOT, 'examples/weboffice-workbench'),
];

const IGNORE_DIRS = new Set(['node_modules', '.vite']);
const violations = [];
const fontInventory = [];

function classifyFontLicense(relPath) {
  if (relPath.includes('/fandol/')) return 'GPL-3.0-with-font-exception (Fandol 0.3)';
  if (relPath.includes('/wps/')) return 'UNAUTHORIZED_WPS_DESKTOP';
  if (relPath.includes('/dist/assets/')) return '@docx-editor.dev/fonts bundled (OFL-1.1 / GUST / MIT)';
  return 'CUSTOM_OFFICIAL_FONT (verify customer web license)';
}

function walk(dir, inDist = false) {
  if (!existsSync(dir)) return;
  for (const name of readdirSync(dir)) {
    if (IGNORE_DIRS.has(name)) continue;
    const full = join(dir, name);
    const st = statSync(full);
    if (st.isDirectory()) {
      walk(full, inDist || name === 'dist');
      continue;
    }
    if (/\.(ttf|otf|woff2?)$/i.test(name)) {
      const rel = full.slice(ROOT.length + 1);
      const lic = classifyFontLicense(rel);
      fontInventory.push({ rel, bytes: st.size, license: lic });
      if (lic === 'UNAUTHORIZED_WPS_DESKTOP') {
        violations.push(`${rel}: unauthorized WPS desktop font bundled`);
      }
      continue;
    }
    if (!/\.(json|ts|tsx|js|mjs|cjs|vue|html|css|map)$/i.test(name)) continue;
    const text = readFileSync(full, 'utf8');
    const isJson = name.endsWith('.json');
    const isMap = name.endsWith('.map');
    if (isMap) {
      for (const forbiddenDir of ['packages/pro/', 'packages/editor-api/', 'packages/docx-to-pdf/', 'packages/telemetry/']) {
        if (text.includes(forbiddenDir)) {
          violations.push(`${full}: source map bundles forbidden Pro source "${forbiddenDir}"`);
        }
      }
      continue;
    }
    if (inDist && text.includes('EigenPal Pro License')) {
      violations.push(`${full}: built artifact contains "EigenPal Pro License" header`);
    }
    for (const pkg of FORBIDDEN_PACKAGES) {
      if (isJson) {
        if (text.includes(`"${pkg}"`)) {
          violations.push(`${full}: declares forbidden Pro dependency "${pkg}"`);
        }
      } else {
        const importRe = new RegExp(`(?:from|import|require)\\s*\\(?\\s*['"]${pkg.replace('/', '\\/')}`, 'g');
        if (importRe.test(text)) {
          violations.push(`${full}: imports forbidden Pro package "${pkg}"`);
        }
      }
    }
  }
}

for (const root of SCAN_ROOTS) {
  walk(root);
}

const wpsFontDir = join(ROOT, 'examples/weboffice-workbench/public/fonts/wps');
if (existsSync(wpsFontDir)) {
  violations.push(`${wpsFontDir}: unauthorized WPS desktop font directory must not be bundled`);
}

if (violations.length > 0) {
  console.error('❌ Apache-2.0 / 字体合规门禁失败：');
  for (const v of violations) console.error(`  - ${v}`);
  process.exit(1);
}

console.log('✅ Apache-2.0 合规门禁通过：源码与 dist 产物均零 Pro 专有包引用，零未授权 WPS 字体打包。');
if (fontInventory.length > 0) {
  console.log(`📦 随包字体清单（共 ${fontInventory.length} 项，另含 @docx-editor.dev/fonts 内置 OFL-1.1 / GUST 西文字体）：`);
  for (const item of fontInventory) {
    console.log(`   - ${item.rel} (${item.bytes} bytes) — ${item.license}`);
  }
}
