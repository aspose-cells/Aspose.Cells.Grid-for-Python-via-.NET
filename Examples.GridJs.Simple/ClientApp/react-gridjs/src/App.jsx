import { useCallback, useEffect, useState } from 'react';
import { GridJsSpreadsheet } from 'gridjs-spreadsheet/react';

const workbookName = 'sample.xlsx';

function createUid() {
  return globalThis.crypto?.randomUUID?.() || `react-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function launcherUrl() {
  return import.meta.env.DEV ? 'http://127.0.0.1:2022/' : '/';
}

function openFileUrl() {
  return import.meta.env.DEV ? 'http://127.0.0.1:2022/legacy' : '/legacy';
}

async function fetchWorkbook(signal) {
  const query = new URLSearchParams({ filename: workbookName, uid: createUid() });
  const response = await fetch(`/GridJs/LoadSpreadsheet?${query}`, { signal });
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  return JSON.parse(await response.text());
}

export function App() {
  const [workbook, setWorkbook] = useState(null);
  const [error, setError] = useState('');
  const [requestId, setRequestId] = useState(0);

  useEffect(() => {
    const controller = new AbortController();
    setWorkbook(null);
    setError('');
    fetchWorkbook(controller.signal)
      .then(setWorkbook)
      .catch((reason) => {
        if (reason.name !== 'AbortError') setError(reason.message || String(reason));
      });
    return () => controller.abort();
  }, [requestId]);

  const handleReady = useCallback((instance, adapter) => {
    if (adapter.payload?.actname) instance.setActiveSheetByName?.(adapter.payload.actname);
    instance.setActiveCell?.(adapter.payload?.actrow || 0, adapter.payload?.actcol || 0);
    instance.setOpenFileUrl?.(openFileUrl());
    console.log('[GridJS React] ready', adapter.payload?.filename);
  }, []);

  return (
    <main className="editor-shell">
      <header className="demo-bar">
        <a className="back-link" href={launcherUrl()} aria-label="Back to framework examples">← Examples</a>
        <div className="demo-title">
          <span className="framework-mark" aria-hidden="true">R</span>
          <span><strong>React</strong><small>{workbookName}</small></span>
        </div>
        <span className="npm-label">npm component</span>
      </header>

      <section className="grid-stage" aria-live="polite">
        {!workbook && !error && (
          <div className="state-panel loading-state" role="status">
            <span className="loading-line wide" />
            <span className="loading-line" />
            <p>Loading workbook from Flask…</p>
          </div>
        )}
        {error && (
          <div className="state-panel error-state" role="alert">
            <strong>Workbook could not be loaded</strong>
            <p>{error}</p>
            <button type="button" onClick={() => setRequestId((value) => value + 1)}>Retry loading</button>
          </div>
        )}
        {workbook && (
          <GridJsSpreadsheet
            data={workbook}
            height="calc(100dvh - 64px)"
            mode="edit"
            locale="en"
            showToolbar
            showContextmenu
            onReady={handleReady}
            onChange={() => console.log('[GridJS React] changed')}
            onError={(payload) => console.error('[GridJS React]', payload)}
          />
        )}
      </section>
    </main>
  );
}
