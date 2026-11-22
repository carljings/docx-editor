<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import {
  WebOfficeDocxEditor,
  probeBrowserEnvironment,
  type FontDiagnosticsSnapshot,
  type WebOfficeEditorHandle,
} from '@weboffice/vue';

const DEFAULT_DOC_URL = `${import.meta.env.BASE_URL}日盯办报告-全指标模板.docx`;
const DEFAULT_FILE_NAME = '日盯办报告-全指标模板.docx';

const editorRef = ref<WebOfficeEditorHandle | null>(null);
const docBytes = ref<Uint8Array | undefined>(undefined);
const currentFileName = ref(DEFAULT_FILE_NAME);
const wpsCompat = ref(
  typeof window !== 'undefined' &&
    new URLSearchParams(window.location.search).get('wpsCompat') === '1'
);
const docRenderKey = ref(0);
const statusText = ref('正在加载样稿...');
const pageCount = ref(0);
const firstPageCount = ref<number | null>(null);
const settledPageCount = ref<number | null>(null);
const lastExportSha = ref('');
const fontDiag = ref<FontDiagnosticsSnapshot>({
  status: 'idle',
  pendingCount: 0,
  requestedFamilies: [],
  hits: [],
  misses: [],
});
const browserEnv = probeBrowserEnvironment();

const fontOptions = {
  baseUrl: import.meta.env.BASE_URL,
  enableFandolFallback: true,
  enableWpsLocalStandIns: false,
  enablePackagedFonts: true,
};

let pageTimer: ReturnType<typeof setInterval> | undefined;

const fontStatusLabel = computed(() => {
  if (fontDiag.value.status === 'settled') return '已稳定';
  if (fontDiag.value.status === 'loading') return `解析中(${fontDiag.value.pendingCount})`;
  return '待触发';
});

const hitSummary = computed(() => {
  if (!fontDiag.value.hits.length)
    return fontDiag.value.status === 'settled' ? '无外部字体命中' : '加载中';
  return fontDiag.value.hits
    .map((h) =>
      h.substitutedFor?.length
        ? `${h.family}(替:${h.substitutedFor[0]})`
        : `${h.family}[${h.layer}]`
    )
    .join(' · ');
});

function toggleWpsCompat() {
  wpsCompat.value = !wpsCompat.value;
  firstPageCount.value = null;
  settledPageCount.value = null;
  docRenderKey.value += 1;
  statusText.value = wpsCompat.value
    ? '已开启 WPS 排版兼容（左对齐中文标点压缩）并重新排版'
    : '已切回 MS Word 默认排版并重新排版';
  publishDiag();
}

function publishDiag() {
  (window as unknown as Record<string, unknown>).__WEBOFFICE_DIAG__ = {
    fileName: currentFileName.value,
    wpsCompat: wpsCompat.value,
    pageCount: pageCount.value,
    firstPageCount: firstPageCount.value,
    settledPageCount: settledPageCount.value,
    browser: browserEnv,
    fonts: fontDiag.value,
    lastExportSha: lastExportSha.value,
  };
}

function onFontDiagnostics(snap: FontDiagnosticsSnapshot) {
  fontDiag.value = snap;
  if (snap.status === 'settled') {
    const count = editorRef.value?.getPageCount() ?? pageCount.value;
    if (count > 0) settledPageCount.value = count;
  }
  publishDiag();
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) return '(需安全上下文)';
  const digest = await crypto.subtle.digest('SHA-256', bytes.slice().buffer);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

async function loadDefaultTemplate() {
  statusText.value = '正在载入《日盯办报告-全指标模板.docx》...';
  const resp = await fetch(DEFAULT_DOC_URL);
  if (!resp.ok) {
    statusText.value = `默认样稿加载失败: HTTP ${resp.status}`;
    return;
  }
  const buf = await resp.arrayBuffer();
  firstPageCount.value = null;
  settledPageCount.value = null;
  currentFileName.value = DEFAULT_FILE_NAME;
  docBytes.value = new Uint8Array(buf);
  statusText.value = `已加载 ${DEFAULT_FILE_NAME} (${buf.byteLength} 字节)`;
}

async function onPickFile(event: Event) {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (!file) return;
  const buf = await file.arrayBuffer();
  firstPageCount.value = null;
  settledPageCount.value = null;
  currentFileName.value = file.name;
  docBytes.value = new Uint8Array(buf);
  statusText.value = `已打开本地文件 ${file.name} (${buf.byteLength} 字节)`;
}

function insertSampleMetric() {
  const res = editorRef.value?.insertTextAtCursor('【指标：接报警情总数】', {
    allowFallback: true,
  });
  const posNote = res?.position ? ` (${res.position.blockId}@${res.position.offset})` : '';
  if (res?.ok && res.changed) {
    statusText.value = res.fallbackApplied
      ? `未定位光标，已兜底插入至首段末尾：【指标：接报警情总数】${posNote}`
      : `已在光标处插入：【指标：接报警情总数】${posNote}`;
  } else {
    statusText.value = `插入失败：${res?.reason ?? '未知错误'}`;
  }
}

async function exportDocx() {
  const bytes = await editorRef.value?.save();
  if (!bytes) {
    statusText.value = '导出失败：编辑器尚未就绪';
    return;
  }
  lastExportSha.value = await sha256Hex(bytes);
  const blob = new Blob([bytes.slice().buffer], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = currentFileName.value.replace(/\.docx$/i, '') + '-weboffice导出.docx';
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  statusText.value = `导出成功 (${bytes.byteLength} 字节)，SHA-256: ${lastExportSha.value.slice(0, 12)}...`;
}

onMounted(() => {
  void loadDefaultTemplate();
  publishDiag();
  pageTimer = setInterval(() => {
    const count = editorRef.value?.getPageCount() ?? 0;
    if (count > 0 && firstPageCount.value === null) firstPageCount.value = count;
    if (count !== pageCount.value) pageCount.value = count;
    if (fontDiag.value.status === 'settled' && count > 0) settledPageCount.value = count;
    publishDiag();
  }, 300);
});

onBeforeUnmount(() => {
  if (pageTimer) clearInterval(pageTimer);
});
</script>

<template>
  <div class="wb-shell">
    <header class="wb-bar">
      <div class="wb-brand">
        <strong>WebOffice 验收工作台</strong>
        <span class="wb-badge">Apache-2.0 纯净核</span>
        <span class="wb-meta">
          页数：<b data-testid="page-count">{{ pageCount }}</b>
          <span v-if="firstPageCount !== null" data-testid="page-milestones">
            （首次 {{ firstPageCount }} / 稳定 {{ settledPageCount ?? '待稳定' }}）
          </span>
        </span>
        <span class="wb-meta"
          >内核：Chromium {{ browserEnv.chromiumMajor ?? '未知' }} · WASM
          {{ browserEnv.hasWasm ? '✓' : '✗' }}</span
        >
      </div>
      <div class="wb-actions">
        <button
          type="button"
          class="wb-btn"
          data-testid="toggle-wps-compat"
          @click="toggleWpsCompat"
        >
          排版基准：{{ wpsCompat ? 'WPS 兼容' : 'Word 默认' }}
        </button>
        <label class="wb-btn">
          打开本地 DOCX
          <input type="file" accept=".docx" hidden @change="onPickFile" />
        </label>
        <button
          type="button"
          class="wb-btn"
          data-testid="insert-metric"
          @click="insertSampleMetric"
        >
          光标插入测试指标
        </button>
        <button
          type="button"
          class="wb-btn wb-btn--primary"
          data-testid="export-docx"
          @click="exportDocx"
        >
          导出原生 DOCX
        </button>
      </div>
    </header>

    <div class="wb-diag">
      <span
        ><strong>状态：</strong><span data-testid="status-text">{{ statusText }}</span></span
      >
      <span
        ><strong>字体状态：</strong
        ><span data-testid="font-status">{{ fontStatusLabel }}</span></span
      >
      <span
        ><strong>字体命中：</strong><span data-testid="font-hits">{{ hitSummary }}</span></span
      >
      <span v-if="lastExportSha"
        ><strong>导出 SHA-256：</strong><code>{{ lastExportSha.slice(0, 16) }}</code></span
      >
    </div>

    <main class="wb-stage">
      <WebOfficeDocxEditor
        v-if="docBytes"
        :key="docRenderKey"
        ref="editorRef"
        :document="docBytes"
        :title="currentFileName"
        :font-options="fontOptions"
        :on-font-diagnostics="onFontDiagnostics"
        :wps-compat="wpsCompat"
        mode="edit"
      />
    </main>
  </div>
</template>

<style>
* {
  box-sizing: border-box;
}
html,
body,
#app {
  margin: 0;
  padding: 0;
  width: 100%;
  height: 100%;
  font-family: -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif;
}
.wb-shell {
  width: 100%;
  height: 100%;
  display: flex;
  flex-direction: column;
  background: #f3f4f6;
}
.wb-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 8px 14px;
  background: #0f172a;
  color: #f8fafc;
  font-size: 13px;
}
.wb-brand {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}
.wb-badge {
  padding: 2px 8px;
  border-radius: 999px;
  background: #16a34a;
  color: #fff;
  font-size: 11px;
  font-weight: 600;
}
.wb-meta {
  color: #cbd5e1;
  font-size: 12px;
}
.wb-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}
.wb-btn {
  padding: 5px 11px;
  border: 1px solid #334155;
  border-radius: 6px;
  background: #1e293b;
  color: #f8fafc;
  font-size: 12px;
  cursor: pointer;
}
.wb-btn:hover {
  background: #334155;
}
.wb-btn--primary {
  border-color: #2563eb;
  background: #2563eb;
}
.wb-btn--primary:hover {
  background: #1d4ed8;
}
.wb-diag {
  display: flex;
  align-items: center;
  gap: 16px;
  flex-wrap: wrap;
  padding: 6px 14px;
  border-bottom: 1px solid #e2e8f0;
  background: #ffffff;
  color: #334155;
  font-size: 12px;
}
.wb-stage {
  flex: 1;
  min-height: 0;
  display: flex;
}
</style>
