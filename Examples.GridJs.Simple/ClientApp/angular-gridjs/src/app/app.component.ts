import { CommonModule } from '@angular/common';
import { Component, OnInit, signal } from '@angular/core';
import { GridJsSpreadsheetComponent } from 'gridjs-spreadsheet/angular';

type GridJsReadyEvent = {
  instance: any;
  adapter: { payload?: Record<string, any> };
};

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, GridJsSpreadsheetComponent],
  templateUrl: './app.component.html',
})
export class AppComponent implements OnInit {
  readonly workbookName = 'sample.xlsx';
  readonly launcherUrl = window.location.port === '4200' ? 'http://127.0.0.1:2022/' : '/';
  readonly openFileUrl = window.location.port === '4200'
    ? 'http://127.0.0.1:2022/legacy'
    : '/legacy';
  readonly workbook = signal<Record<string, any> | null>(null);
  readonly error = signal('');

  ngOnInit(): void {
    void this.loadWorkbook();
  }

  async loadWorkbook(): Promise<void> {
    this.workbook.set(null);
    this.error.set('');
    try {
      const uid = globalThis.crypto?.randomUUID?.() || `angular-${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const query = new URLSearchParams({ filename: this.workbookName, uid });
      const response = await fetch(`/GridJs/LoadSpreadsheet?${query}`);
      if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
      this.workbook.set(JSON.parse(await response.text()) as Record<string, any>);
    } catch (reason: unknown) {
      this.error.set(reason instanceof Error ? reason.message : String(reason));
    }
  }

  handleReady(event: GridJsReadyEvent): void {
    const payload = event.adapter.payload;
    if (payload?.['actname']) event.instance.setActiveSheetByName?.(payload['actname']);
    event.instance.setActiveCell?.(payload?.['actrow'] || 0, payload?.['actcol'] || 0);
    event.instance.setOpenFileUrl?.(this.openFileUrl);
    console.log('[GridJS Angular] ready', payload?.['filename']);
  }

  handleError(payload: unknown): void {
    console.error('[GridJS Angular]', payload);
  }

  logChange(): void {
    console.log('[GridJS Angular] changed');
  }
}
