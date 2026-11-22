import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import path from 'node:path';

const monorepoRoot = path.resolve(__dirname, '../..');

export default defineConfig({
  base: process.env.VITE_BASE_PATH ?? '/',
  plugins: [vue()],
  root: __dirname,
  server: {
    host: '127.0.0.1',
    port: 5188,
    strictPort: false,
  },
  build: {
    // 面向国产化老内核浏览器（奇安信 / 360 信创 / 统信 / 红莲花常见 Chromium 86+）
    target: 'chrome86',
    sourcemap: false,
  },
  resolve: {
    alias: [
      {
        find: /^@weboffice\/vue\/styles\.css$/,
        replacement: path.join(monorepoRoot, 'packages/weboffice-vue/dist/styles.css'),
      },
      {
        find: /^@weboffice\/vue\/polyfills$/,
        replacement: path.join(monorepoRoot, 'packages/weboffice-vue/src/polyfills.ts'),
      },
      {
        find: /^@weboffice\/vue$/,
        replacement: path.join(monorepoRoot, 'packages/weboffice-vue/src/index.ts'),
      },
    ],
  },
});
