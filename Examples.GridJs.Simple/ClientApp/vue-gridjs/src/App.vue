<script setup>
import { onBeforeUnmount, onMounted, ref } from 'vue';
import { GridJsSpreadsheet } from 'gridjs-spreadsheet/vue';

const workbookName = 'sample.xlsx';
const workbook = ref(null);
const error = ref('');
let controller = null;

function createUid() {
  return globalThis.crypto?.randomUUID?.() || `vue-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function launcherUrl() {
  return import.meta.env.DEV ? 'http://127.0.0.1:2022/' : '/';
}

function openFileUrl() {
  return import.meta.env.DEV ? 'http://127.0.0.1:2022/legacy' : '/legacy';
}

async function loadWorkbook() {
  controller?.abort();
  controller = new AbortController();
  workbook.value = null;
  error.value = '';
  try {
    const query = new URLSearchParams({ filename: workbookName, uid: createUid() });
    const response = await fetch(`/GridJs/LoadSpreadsheet?${query}`, { signal: controller.signal });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    workbook.value = JSON.parse(await response.text());
  } catch (reason) {
    if (reason.name !== 'AbortError') error.value = reason.message || String(reason);
  }
}

function handleReady(instance, adapter) {
  if (adapter.payload?.actname) instance.setActiveSheetByName?.(adapter.payload.actname);
  instance.setActiveCell?.(adapter.payload?.actrow || 0, adapter.payload?.actcol || 0);
  instance.setOpenFileUrl?.(openFileUrl());
  console.log('[GridJS Vue] ready', adapter.payload?.filename);
}

onMounted(loadWorkbook);
onBeforeUnmount(() => controller?.abort());
</script>

<template>
  <main class="editor-shell">
    <header class="demo-bar">
      <a class="back-link" :href="launcherUrl()" aria-label="Back to framework examples">← Examples</a>
      <div class="demo-title">
        <span class="framework-mark" aria-hidden="true">V</span>
        <span><strong>Vue 3</strong><small>{{ workbookName }}</small></span>
      </div>
      <span class="npm-label">npm component</span>
    </header>

    <section class="grid-stage" aria-live="polite">
      <div v-if="!workbook && !error" class="state-panel loading-state" role="status">
        <span class="loading-line wide"></span>
        <span class="loading-line"></span>
        <p>Loading workbook from Flask…</p>
      </div>
      <div v-else-if="error" class="state-panel error-state" role="alert">
        <strong>Workbook could not be loaded</strong>
        <p>{{ error }}</p>
        <button type="button" @click="loadWorkbook">Retry loading</button>
      </div>
      <GridJsSpreadsheet
        v-else
        :data="workbook"
        height="calc(100dvh - 64px)"
        mode="edit"
        locale="en"
        :show-toolbar="true"
        :show-contextmenu="true"
        @ready="handleReady"
        @change="console.log('[GridJS Vue] changed')"
        @error="payload => console.error('[GridJS Vue]', payload)"
      />
    </section>
  </main>
</template>
