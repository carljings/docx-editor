import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    index: 'src/index.ts',
    polyfills: 'src/polyfills.ts',
  },
  platform: 'browser',
  target: 'chrome86',
  format: ['esm', 'cjs'],
  dts: true,
  clean: true,
  splitting: false,
  sourcemap: true,
  external: [
    'vue',
    '@docx-editor.dev/core',
    '@docx-editor.dev/core/editor',
    '@docx-editor.dev/core/layout',
    '@docx-editor.dev/core/contracts/editor',
    '@docx-editor.dev/vue',
    '@docx-editor.dev/i18n',
    '@docx-editor.dev/fonts',
    '@docx-editor.dev/fonts-cjk',
  ],
});
