import Spreadsheet from 'gridjs-spreadsheet';
import JSZip from 'jszip';
import 'gridjs-spreadsheet/xspreadsheet.css';
import './styles.css';

const workbookName = 'sample.xlsx';
const launcher = import.meta.env.DEV ? 'http://127.0.0.1:2022/' : '/';
const openFileUrl = import.meta.env.DEV ? 'http://127.0.0.1:2022/legacy' : '/legacy';
const host = document.getElementById('spreadsheet');
const loadingState = document.getElementById('loading-state');
const errorState = document.getElementById('error-state');
const errorMessage = document.getElementById('error-message');
let spreadsheet = null;

window.JSZip = JSZip;
document.getElementById('back-link').href = launcher;
document.getElementById('retry-button').addEventListener('click', loadWorkbook);
window.addEventListener('beforeunload', () => spreadsheet?.destroy?.());

function createUid() {
  return globalThis.crypto?.randomUUID?.() || `npm-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function loadWorkbook() {
  spreadsheet?.destroy?.();
  spreadsheet = null;
  host.innerHTML = '';
  loadingState.hidden = false;
  errorState.hidden = true;

  try {
    const query = new URLSearchParams({ filename: workbookName, uid: createUid() });
    const response = await fetch(`/GridJs/LoadSpreadsheet?${query}`);
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const payload = JSON.parse(await response.text());

    spreadsheet = new Spreadsheet(host, {
      updateMode: 'server',
      updateUrl: '/GridJs/UpdateCell',
      mode: 'edit',
      local: 'en',
      showToolbar: true,
      showContextmenu: true,
    });
    spreadsheet.loadData(payload.data || [], payload.actname);
    spreadsheet.setUniqueId?.(payload.uniqueid || '');
    spreadsheet.setFileName?.(payload.filename || workbookName);
    spreadsheet.setImageInfo?.('/GridJs/ImageUrl', '/GridJs/AddImage', '/GridJs/AddImageByURL', '/GridJs/CopyImage', 5678, '/content/img/loading.gif');
    spreadsheet.setFileDownloadInfo?.('/GridJs/Download');
    spreadsheet.setOleDownloadInfo?.('/GridJs/Ole');
    spreadsheet.setLazyLoadingUrl?.('/GridJs/LazyLoadingStreamJson');
    if (payload.actname) spreadsheet.setActiveSheetByName?.(payload.actname);
    spreadsheet.setActiveCell?.(payload.actrow || 0, payload.actcol || 0);
    spreadsheet.setOpenFileUrl?.(openFileUrl);
    spreadsheet.updateCellError?.((message) => console.error('[GridJS npm]', message));
    loadingState.hidden = true;
    console.log('[GridJS npm] ready', payload.filename);
  } catch (reason) {
    loadingState.hidden = true;
    errorMessage.textContent = reason.message || String(reason);
    errorState.hidden = false;
  }
}

loadWorkbook();
