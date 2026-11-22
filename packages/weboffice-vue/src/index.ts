/**
 * @weboffice/vue — 纯 Apache-2.0 WebOffice Vue 3 封装层
 * 严格禁止引入 @docx-editor.dev/pro、editor-api、docx-to-pdf、telemetry 四个专有许可包。
 */

export {
  installLegacyPolyfills,
  probeBrowserEnvironment,
  type BrowserEnvironmentProbe,
  type PolyfillReport,
} from './polyfills';

export {
  CN_FONT_SIZES,
  CN_OFFICIAL_FONT_FILES,
  FANDOL_STAND_INS,
  WPS_LOCAL_DEV_STAND_INS,
  createOfficialFontResolver,
  type ChineseFontSizeOption,
  type FontDiagnosticsSnapshot,
  type FontDiagnosticsStatus,
  type FontHitRecord,
  type FontMissRecord,
  type FontStandInSpec,
  type OfficialFontFileSpec,
  type OfficialFontResolverOptions,
} from './fonts';

export {
  createWebOfficeAdapter,
  type CursorPosition,
  type InsertTextMode,
  type InsertTextOptions,
  type InsertTextResult,
  type WebOfficeDiagnostics,
  type WebOfficeEditorHandle,
} from './adapter';

export { ChineseFontSize } from './ChineseFontSize';
export { WebOfficeDocxEditor } from './WebOfficeDocxEditor';

// 安全透传纯 Apache-2.0 上游核心与 Vue 能力，方便报告中心单包接入
export {
  DocxEditor,
  useDocxEditor,
  useEditorCommand,
  useEditorValueCommand,
  useFonts,
  useHighlights,
  type DocxEditorProps,
  type DocxEditorRef,
  type EditorMode,
} from '@docx-editor.dev/vue';
export type {
  DocumentHandle,
  DocumentSource,
  Editor,
  EditorCommand,
  HighlightRange,
} from '@docx-editor.dev/core/contracts/editor';
export {
  composeFontOrigins,
  customFonts,
  defineFontResolver,
  type FontOrigin,
  type FontUrlSource,
} from '@docx-editor.dev/core/editor';
export {
  getWpsCompatMode,
  setHarfBuzzWasmUrl,
  setWpsCompatMode,
  type WpsCompatOptions,
} from '@docx-editor.dev/core/layout';
export { packagedFonts } from '@docx-editor.dev/fonts';
export { zhCN, type Translations } from '@docx-editor.dev/i18n';
