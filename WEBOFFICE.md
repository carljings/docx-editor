# WebOffice 二开说明（基于 EigenPal Apache-2.0 核心）

本仓库为 `eigenpal/docx-editor` 的二次开发 Fork（`carljings/docx-editor`），面向政务/报告中心中文公文场景与国产化信创浏览器，严格只使用 **Apache-2.0** 开源核心包。

## 目录分工（与上游低冲突）

- `packages/core`、`packages/vue`、`packages/i18n`、`packages/fonts`、`packages/fonts-cjk`：上游 Apache-2.0 核心源码（后续按需定向修中文排版、表格跨页与网格细节）。
- `packages/weboffice-vue`（包名 `@weboffice/vue`）：**我方自维护的纯 Apache-2.0 Vue 3 封装层**，包含：
  - `src/polyfills.ts`：国产化老内核浏览器（Chromium 80～98）运行时 Polyfill（`structuredClone`、`Object.hasOwn`、`Array/String/TypedArray.prototype.at`、`replaceAll`）与 `probeBrowserEnvironment()` 探针；
  - `src/fonts.ts`：公文中文字号表 `CN_FONT_SIZES`（初号～八号）与离线公文字体解析器 `createOfficialFontResolver()`（默认禁用无授权本机 WPS 字体）；
  - `src/adapter.ts`：引擎无关的业务适配契约 `WebOfficeEditorHandle`（`load` / `save` / `insertTextAtCursor` / `getPageCount` / `getDiagnostics`）；
  - `src/ChineseFontSize.ts` + `src/WebOfficeDocxEditor.ts`：开箱即用 Vue 3 组件。
- `examples/weboffice-workbench`：纯 Apache-2.0 的独立验收工作台（端口 `5188`，构建目标 `chrome86`，内置《日盯办报告-全指标模板.docx》与开源 Fandol 兜底字体）。
- `scripts/check-weboffice-apache-only.mjs`：合规门禁脚本，禁止引入 `pro`、`editor-api`、`docx-to-pdf`、`telemetry` 及未授权 WPS 桌面字体。

## 常用命令

```bash
# 1. 构建全部 Apache-2.0 核心包 + @weboffice/vue + 跑合规门禁
bun run weboffice:build

# 2. 跑 @weboffice/vue 单测 + 合规检查 + chrome86 生产构建验收
bun run weboffice:test

# 3. 启动本地验收工作台（http://127.0.0.1:5188）
bun run weboffice:dev

# 4. 单独跑 Apache-2.0 合规门禁
bun run weboffice:check-license
```

## 报告中心如何直接引用本工程源码

在报告中心前端工程（例如 `jrdmx/docs/报告中心/eigenpal-docx-redesign` 或正式 Vue 3 工程）中，有两种方式直接引用本 Fork：

### 方式 A：推荐——通过 `package.json` 本地工作区/文件依赖引用 `@weboffice/vue`

```json
{
  "dependencies": {
    "@weboffice/vue": "file:../../../../weboffice/docx-editor/packages/weboffice-vue",
    "@docx-editor.dev/core": "file:../../../../weboffice/docx-editor/packages/core",
    "@docx-editor.dev/vue": "file:../../../../weboffice/docx-editor/packages/vue",
    "@docx-editor.dev/i18n": "file:../../../../weboffice/docx-editor/packages/i18n",
    "@docx-editor.dev/fonts": "file:../../../../weboffice/docx-editor/packages/fonts"
  }
}
```

### 方式 B：Vite 开发态直连源码热更新（`vite.config.ts`）

```ts
import path from 'node:path';

const webofficeRoot = path.resolve(__dirname, '../../../../weboffice/docx-editor');

export default defineConfig({
  build: { target: 'chrome86' },
  resolve: {
    alias: {
      '@weboffice/vue/styles.css': path.join(
        webofficeRoot,
        'packages/weboffice-vue/dist/styles.css'
      ),
      '@weboffice/vue/polyfills': path.join(
        webofficeRoot,
        'packages/weboffice-vue/src/polyfills.ts'
      ),
      '@weboffice/vue': path.join(webofficeRoot, 'packages/weboffice-vue/src/index.ts'),
      '@docx-editor.dev/vue/styles.css': path.join(
        webofficeRoot,
        'packages/vue/src/styles/editor.css'
      ),
      '@docx-editor.dev/vue': path.join(webofficeRoot, 'packages/vue/dist/index.js'),
      '@docx-editor.dev/i18n': path.join(webofficeRoot, 'packages/i18n/dist/index.mjs'),
      '@docx-editor.dev/fonts': path.join(webofficeRoot, 'packages/fonts/dist/index.js'),
      '@docx-editor.dev/core/editor': path.join(webofficeRoot, 'packages/core/dist/editor.js'),
      '@docx-editor.dev/core/layout': path.join(webofficeRoot, 'packages/core/dist/layout.js'),
      '@docx-editor.dev/core': path.join(webofficeRoot, 'packages/core/dist/index.js'),
    },
  },
});
```

入口文件最顶部先加载 Polyfill 与样式：

```ts
import '@weboffice/vue/polyfills';
import '@weboffice/vue/styles.css';
```

## 同步上游更新

```bash
git fetch upstream
git checkout main && git merge upstream/main
git checkout dev/weboffice && git merge main
bun run weboffice:build && bun run weboffice:test
```
