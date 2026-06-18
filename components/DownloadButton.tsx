'use client';

import type { CompareResult, DiffRow } from '@/lib/compare';
import * as XLSX from 'xlsx';

type DownloadButtonProps = {
  result: CompareResult | null;
};

const FILL = {
  headerBg:  'D1D5DB',
  skuChanged:'FCD34D',
  skuMissing:'FCA5A5',
  skuAdded:  '6EE7B7',
  aifyRow:   'FFF7ED',
  aifyCell:  'FED7AA',
  ambraRow:  'F0FDF4',
  ambraCell: 'BBF7D0',
  missingRow:'FEE2E2',
  addedRow:  'D1FAE5',
  matched:   'FFFFFF',
} as const;

function writeCell(
  ws: XLSX.WorkSheet,
  r: number,
  c: number,
  value: string | number,
  fillRgb?: string,
  bold?: boolean,
  fontSize?: number,
): void {
  const ref = XLSX.utils.encode_cell({ r, c });
  ws[ref] = typeof value === 'number' ? { t: 'n', v: value } : { t: 's', v: value };

  const hasFill = !!fillRgb;
  const hasFont = bold || !!fontSize;

  if (hasFill || hasFont) {
    ws[ref].s = {
      // bgColor: indexed 64 is required by the OOXML spec alongside fgColor
      // for solid fills to actually render in Excel.
      ...(hasFill
        ? { fill: { patternType: 'solid', fgColor: { rgb: fillRgb }, bgColor: { indexed: 64 } } }
        : {}),
      ...(hasFont ? { font: { bold: !!bold, sz: fontSize ?? 11 } } : {}),
    };
  }
}

function setRef(ws: XLSX.WorkSheet, lastRow: number, lastCol: number): void {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: lastCol } });
}

function buildSummarySheet(
  stats: CompareResult['stats'],
  columnChangeCounts: CompareResult['columnChangeCounts'],
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};
  let row = 0;

  // ── Title ──────────────────────────────────────────────────────────────────
  writeCell(ws, row, 0, 'Akeneo QA Tool — Comparison Report', '3B5BDB', true, 14);
  writeCell(ws, row, 1, '', '3B5BDB');
  row += 2;

  // ── Statistics ─────────────────────────────────────────────────────────────
  writeCell(ws, row, 0, 'STATISTICS', FILL.headerBg, true);
  writeCell(ws, row, 1, 'COUNT',      FILL.headerBg, true);
  writeCell(ws, row, 2, 'STATUS',     FILL.headerBg, true);
  row++;

  const statRows: [string, number, string, string][] = [
    ['Total SKUs',               stats.total,   'DBEAFE', 'All unique SKUs across both files'],
    ['Matched (identical)',      stats.matched, 'F3F4F6', 'All values match — no action needed'],
    ['Changed (values differ)',  stats.changed, 'FEF3C7', '⚠ Values differ between Aify and Ambra'],
    ['Missing in Aify',         stats.missing, 'FEE2E2', '✗ SKU exists in Ambra but is missing from Aify'],
    ['Aify Only (not in Ambra)', stats.added,   'D1FAE5', '+ SKU is in Aify but does not exist in Ambra'],
  ];

  for (const [label, count, fill, note] of statRows) {
    writeCell(ws, row, 0, label, fill, true);
    writeCell(ws, row, 1, count, fill);
    writeCell(ws, row, 2, note,  fill);
    row++;
  }
  row++;

  // ── Colour legend ──────────────────────────────────────────────────────────
  writeCell(ws, row, 0, 'COLOUR LEGEND', FILL.headerBg, true);
  writeCell(ws, row, 1, 'MEANING',       FILL.headerBg, true);
  writeCell(ws, row, 2, 'ACTION',        FILL.headerBg, true);
  row++;

  const legendRows: [string, string, string, string][] = [
    [FILL.aifyCell,   'AIFY row (orange)',        'Your current Aify file value',          'Review — this may need updating'],
    [FILL.ambraCell,  'AMBRA row (green)',         'Ambra source-of-truth value',           'This is the correct target value'],
    ['FCA5A5',        'Changed cell — Aify',       'Old value in your file',                'Replace with the Ambra value'],
    [FILL.ambraCell,  'Changed cell — Ambra',      'Correct value from Ambra',              'Use this value in your file'],
    [FILL.missingRow, 'Missing row (light red)',   'SKU in Ambra only — not in Aify',       'Add this SKU to your Aify file'],
    [FILL.addedRow,   'Aify Only row (light green)','SKU in Aify only — not in Ambra',      'Verify whether this SKU should exist'],
    [FILL.skuChanged, 'SKU cell (amber)',           'This SKU has at least one change',      'Scan Differences sheet for this SKU'],
  ];

  for (const [fill, label, meaning, action] of legendRows) {
    writeCell(ws, row, 0, label,   fill, true);
    writeCell(ws, row, 1, meaning, fill);
    writeCell(ws, row, 2, action,  fill);
    row++;
  }
  row++;

  // ── Top changed columns ────────────────────────────────────────────────────
  const changedCols = Object.entries(columnChangeCounts)
    .filter(([, count]) => count > 0)
    .sort(([, a], [, b]) => b - a)
    .slice(0, 20);

  if (changedCols.length > 0) {
    writeCell(ws, row, 0, 'TOP CHANGED COLUMNS', FILL.headerBg, true);
    writeCell(ws, row, 1, 'CHANGES',             FILL.headerBg, true);
    row++;

    for (const [col, count] of changedCols) {
      writeCell(ws, row, 0, col,   'FEF3C7');
      writeCell(ws, row, 1, count, 'FEF3C7');
      row++;
    }
  }

  setRef(ws, row, 2);
  ws['!cols'] = [{ wch: 36 }, { wch: 12 }, { wch: 52 }];
  return ws;
}

function buildDataSheet(inputRows: DiffRow[], columns: string[]): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};
  const headers = ['Source', 'SKU', ...columns];

  headers.forEach((h, c) => writeCell(ws, 0, c, h, FILL.headerBg, true));

  let rowIdx = 1;

  for (const row of inputRows) {
    if (row.rowStatus === 'changed') {
      // Aify row — what you have
      writeCell(ws, rowIdx, 0, 'AIFY', FILL.aifyCell, true);
      writeCell(ws, rowIdx, 1, row.sku, FILL.skuChanged, true);
      columns.forEach((col, offset) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        writeCell(ws, rowIdx, offset + 2, c.aify, c.status === 'changed' ? FILL.aifyCell : FILL.aifyRow);
      });
      rowIdx++;

      // Ambra row — source of truth
      writeCell(ws, rowIdx, 0, 'AMBRA ✓', FILL.ambraCell, true);
      writeCell(ws, rowIdx, 1, row.sku, FILL.skuChanged, true);
      columns.forEach((col, offset) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        writeCell(ws, rowIdx, offset + 2, c.ambra, c.status === 'changed' ? FILL.ambraCell : FILL.ambraRow, c.status === 'changed');
      });
      rowIdx++;
    } else if (row.rowStatus === 'missing') {
      writeCell(ws, rowIdx, 0, 'MISSING',   FILL.missingRow, true);
      writeCell(ws, rowIdx, 1, row.sku,      FILL.skuMissing, true);
      columns.forEach((col, offset) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'missing' as const };
        writeCell(ws, rowIdx, offset + 2, c.ambra, FILL.missingRow);
      });
      rowIdx++;
    } else if (row.rowStatus === 'added') {
      writeCell(ws, rowIdx, 0, 'AIFY ONLY', FILL.addedRow, true);
      writeCell(ws, rowIdx, 1, row.sku,      FILL.skuAdded, true);
      columns.forEach((col, offset) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'added' as const };
        writeCell(ws, rowIdx, offset + 2, c.aify, FILL.addedRow);
      });
      rowIdx++;
    } else {
      writeCell(ws, rowIdx, 0, 'MATCH', FILL.matched);
      writeCell(ws, rowIdx, 1, row.sku,  FILL.matched, true);
      columns.forEach((col, offset) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        writeCell(ws, rowIdx, offset + 2, c.ambra || c.aify, FILL.matched);
      });
      rowIdx++;
    }
  }

  setRef(ws, Math.max(rowIdx - 1, 1), headers.length - 1);
  ws['!cols'] = [{ wch: 12 }, { wch: 20 }, ...columns.map(() => ({ wch: 22 }))];
  return ws;
}

function buildWorkbook(result: CompareResult): XLSX.WorkBook {
  const workbook = XLSX.utils.book_new();
  const diffRows = result.rows.filter((r) => r.rowStatus !== 'match');

  XLSX.utils.book_append_sheet(workbook, buildSummarySheet(result.stats, result.columnChangeCounts), 'Summary');
  XLSX.utils.book_append_sheet(workbook, buildDataSheet(diffRows, result.columns), 'Differences');
  XLSX.utils.book_append_sheet(workbook, buildDataSheet(result.rows, result.columns), 'Full Comparison');

  return workbook;
}

export default function DownloadButton({ result }: DownloadButtonProps) {
  const handleDownload = () => {
    if (!result) return;
    // bookType:'xls' (BIFF8) is required — SheetJS community edition only
    // writes cell fill/font styles for XLS, XLSB, and XLML formats.
    // The .xlsx format requires SheetJS Pro for style support.
    XLSX.writeFile(buildWorkbook(result), 'akeneo-qa-report.xls', { bookType: 'xls' });
  };

  return (
    <button
      type="button"
      onClick={handleDownload}
      disabled={!result}
      className="flex items-center gap-2 rounded-lg bg-brand px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:bg-gray-300 disabled:shadow-none"
    >
      <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5} aria-hidden="true">
        <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
      </svg>
      Download QA Report (.xls)
    </button>
  );
}
