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
  CN_SELECTABLE_FONT_FAMILIES,
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
  CM_PER_INCH,
  CN_OFFICIAL_SYMBOLS,
  GB_T_9704_MARGINS_CM,
  OFFICIAL_PAPER_SIZES,
  TWIPS_PER_CM,
  TWIPS_PER_INCH,
  TWIPS_PER_PT,
  charsToTwips,
  cmToTwips,
  countDocumentWords,
  ptToTwips,
  stepChineseFontSize,
  twipsToChars,
  twipsToCm,
  twipsToPt,
  type DocumentWordStats,
  type OfficialMarginsCm,
  type OfficialPaperSizePreset,
  type OfficialSymbolOption,
} from './units';

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
export { FindReplacePanel } from './FindReplacePanel';
export { OfficialFormatDialog, type OfficialFormatTab } from './OfficialFormatDialog';
export { WpsQuickBar } from './WpsQuickBar';
export { WebOfficeStatusBar } from './WebOfficeStatusBar';
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
