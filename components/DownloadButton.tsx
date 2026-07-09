'use client';

import type { CompareResult, DiffRow } from '@/lib/compare';
import * as XLSX from 'xlsx';

type DownloadButtonProps = {
  result: CompareResult | null;
};

// ─── Colour palette (matches the on-screen Tailwind classes exactly) ──────────
const FILL = {
  // Column header row
  headerBg:      'D1D5DB', // gray-300     — normal column header
  headerChanged: 'FEF3C7', // amber-100    — column that has ≥1 change

  // SKU column accent per row type
  skuChanged:    'FCD34D', // amber-300
  skuMissing:    'FCA5A5', // red-300
  skuAdded:      '6EE7B7', // emerald-300

  // AIFY row  (screen: bg-orange-200 label, bg-white row, bg-red-100 bad cell)
  aifyLabel:     'FED7AA', // orange-200   — "AIFY" badge cell
  aifyRow:       'FFF7ED', // orange-50    — unchanged cell in Aify row
  aifyCellBad:   'FECACA', // red-200      — changed cell (wrong value) ← red like screen

  // AMBRA ✓ row (screen: bg-emerald-300 label, bg-white row, bg-emerald-100 good cell)
  ambraLabel:    '6EE7B7', // emerald-300  — "AMBRA ✓" badge cell
  ambraRow:      'F0FDF4', // green-50     — unchanged cell in Ambra row
  ambraCell:     'A7F3D0', // emerald-200  — changed cell (correct value) ← green like screen

  // Other row types
  missingRow:    'FEE2E2', // red-100
  addedRow:      'D1FAE5', // emerald-100
  matched:       'FFFFFF', // white
} as const;

// ─── Cell writer ──────────────────────────────────────────────────────────────
function writeCell(
  ws: XLSX.WorkSheet,
  r: number,
  c: number,
  value: string | number,
  fillRgb?: string,
  bold?: boolean,
  fontSize?: number,
  wrapText?: boolean,
  fontColor?: string,
): void {
  const ref = XLSX.utils.encode_cell({ r, c });
  ws[ref] = typeof value === 'number' ? { t: 'n', v: value } : { t: 's', v: value };

  const hasFill  = !!fillRgb;
  const hasFont  = bold || !!fontSize || !!fontColor;
  const hasAlign = wrapText !== undefined;

  if (hasFill || hasFont || hasAlign) {
    ws[ref].s = {
      ...(hasFill
        ? { fill: { patternType: 'solid', fgColor: { rgb: fillRgb }, bgColor: { indexed: 64 } } }
        : {}),
      ...(hasFont
        ? { font: { bold: !!bold, sz: fontSize ?? 10, color: fontColor ? { rgb: fontColor } : undefined } }
        : {}),
      ...(hasAlign ? { alignment: { wrapText: !!wrapText, vertical: 'top' } } : {}),
    };
  }
}

function setRef(ws: XLSX.WorkSheet, lastRow: number, lastCol: number): void {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow, c: lastCol } });
}

// ─── Summary sheet ────────────────────────────────────────────────────────────
function buildSummarySheet(
  stats: CompareResult['stats'],
  columnChangeCounts: CompareResult['columnChangeCounts'],
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};
  let row = 0;

  writeCell(ws, row, 0, 'Akeneo QA Tool — Comparison Report', '3B5BDB', true, 14);
  writeCell(ws, row, 1, '', '3B5BDB');
  writeCell(ws, row, 2, '', '3B5BDB');
  row += 2;

  // ── Statistics ──────────────────────────────────────────────────────────────
  writeCell(ws, row, 0, 'STATISTICS',  FILL.headerBg, true, 11);
  writeCell(ws, row, 1, 'COUNT',       FILL.headerBg, true, 11);
  writeCell(ws, row, 2, 'NOTE',        FILL.headerBg, true, 11);
  row++;

  const statRows: [string, number, string, string][] = [
    ['Total SKUs',                stats.total,   'DBEAFE', 'All unique SKUs across both files'],
    ['✔ Matched (identical)',     stats.matched, 'F3F4F6', 'All values match — no action needed'],
    ['⚠ Changed (values differ)', stats.changed, 'FEF3C7', 'Values differ between Aify and Ambra'],
    ['✗ Missing in Aify',        stats.missing, 'FEE2E2', 'SKU exists in Ambra but not in Aify'],
    ['+ Aify Only (not in Ambra)',stats.added,   'D1FAE5', 'SKU in Aify but not in Ambra'],
  ];

  for (const [label, count, fill, note] of statRows) {
    writeCell(ws, row, 0, label, fill, true,  10);
    writeCell(ws, row, 1, count, fill, false, 10);
    writeCell(ws, row, 2, note,  fill, false, 10);
    row++;
  }
  row++;

  // ── Colour legend ────────────────────────────────────────────────────────────
  writeCell(ws, row, 0, 'COLOUR LEGEND',  FILL.headerBg, true, 11);
  writeCell(ws, row, 1, 'MEANING',        FILL.headerBg, true, 11);
  writeCell(ws, row, 2, 'ACTION',         FILL.headerBg, true, 11);
  row++;

  const legendRows: [string, string, string, string][] = [
    [FILL.aifyLabel,   'AIFY row (orange)',           'Your current Aify file value',           'Review — may need updating'],
    [FILL.aifyCellBad, 'Changed cell — AIFY (red)',   'Incorrect value in your file',           'Replace with the green Ambra value'],
    [FILL.ambraLabel,  'AMBRA ✓ row (green)',          'Ambra source-of-truth value',            'This is the correct target value'],
    [FILL.ambraCell,   'Changed cell — AMBRA (green)', 'Correct value from Ambra',               'Copy this into your Aify file'],
    [FILL.missingRow,  'Missing row (light red)',      'SKU in Ambra only — not in Aify',        'Add this SKU to your Aify file'],
    [FILL.addedRow,    'Aify Only row (light green)',  'SKU in Aify only — not in Ambra',        'Verify if this SKU should exist'],
    [FILL.skuChanged,  'SKU cell (amber)',             'This SKU has at least one change',       'Scan Differences sheet for this SKU'],
    [FILL.headerChanged,'Column header (amber)',       'This column has at least one difference','Count shown in brackets e.g. (12)'],
  ];

  for (const [fill, label, meaning, action] of legendRows) {
    writeCell(ws, row, 0, label,   fill, true,  10);
    writeCell(ws, row, 1, meaning, fill, false, 10);
    writeCell(ws, row, 2, action,  fill, false, 10);
    row++;
  }
  row++;

  // ── Top changed columns ──────────────────────────────────────────────────────
  const changedCols = Object.entries(columnChangeCounts)
    .filter(([, n]) => n > 0)
    .sort(([, a], [, b]) => b - a);

  if (changedCols.length > 0) {
    writeCell(ws, row, 0, `TOP CHANGED COLUMNS  (${changedCols.length} columns with differences)`,
      FILL.headerBg, true, 11);
    writeCell(ws, row, 1, 'CHANGES', FILL.headerBg, true, 11);
    writeCell(ws, row, 2, '% OF CHANGED ROWS',  FILL.headerBg, true, 11);
    row++;

    const totalChanged = stats.changed + stats.missing + stats.added;
    for (const [col, count] of changedCols) {
      const pct = totalChanged > 0 ? `${((count / totalChanged) * 100).toFixed(1)}%` : '—';
      writeCell(ws, row, 0, col,   'FEF3C7', false, 10);
      writeCell(ws, row, 1, count, 'FEF3C7', true,  10);
      writeCell(ws, row, 2, pct,   'FEF3C7', false, 10);
      row++;
    }
  }

  setRef(ws, row, 2);
  ws['!cols'] = [{ wch: 40 }, { wch: 12 }, { wch: 52 }];
  return ws;
}

// ─── Data sheet (Differences / Full Comparison) ───────────────────────────────
function buildDataSheet(
  inputRows: DiffRow[],
  columns: string[],
  columnChangeCounts: CompareResult['columnChangeCounts'],
): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};

  // ── Header row — amber for changed columns, shows count in brackets ──────────
  writeCell(ws, 0, 0, 'Source', FILL.headerBg, true, 10);
  writeCell(ws, 0, 1, 'SKU',    FILL.headerBg, true, 10);

  columns.forEach((col, i) => {
    const count = columnChangeCounts[col] ?? 0;
    const label = count > 0 ? `${col}  (${count})` : col;
    const fill  = count > 0 ? FILL.headerChanged : FILL.headerBg;
    writeCell(ws, 0, i + 2, label, fill, true, 10);
  });

  // ── Data rows ─────────────────────────────────────────────────────────────────
  let rowIdx = 1;

  for (const row of inputRows) {

    if (row.rowStatus === 'changed') {
      // ── AIFY row ──
      writeCell(ws, rowIdx, 0, 'AIFY',    FILL.aifyLabel, true,  10, false, '7C2D12');
      writeCell(ws, rowIdx, 1, row.sku,   FILL.skuChanged, true, 10);
      columns.forEach((col, i) => {
        const c    = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        const fill = c.status === 'changed' ? FILL.aifyCellBad : FILL.aifyRow;
        const bold = c.status === 'changed';
        const fc   = c.status === 'changed' ? '991B1B' : undefined; // red-800 text for bad cells
        writeCell(ws, rowIdx, i + 2, c.aify, fill, bold, 10, true, fc);
      });
      rowIdx++;

      // ── AMBRA ✓ row ──
      writeCell(ws, rowIdx, 0, 'AMBRA ✓', FILL.ambraLabel, true,  10, false, '064E3B');
      writeCell(ws, rowIdx, 1, row.sku,   FILL.skuChanged,  true, 10);
      columns.forEach((col, i) => {
        const c    = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        const fill = c.status === 'changed' ? FILL.ambraCell : FILL.ambraRow;
        const bold = c.status === 'changed';
        const fc   = c.status === 'changed' ? '065F46' : undefined; // emerald-900 for good cells
        writeCell(ws, rowIdx, i + 2, c.ambra, fill, bold, 10, true, fc);
      });
      rowIdx++;

    } else if (row.rowStatus === 'missing') {
      writeCell(ws, rowIdx, 0, 'MISSING',  FILL.missingRow, true,  10, false, '7F1D1D');
      writeCell(ws, rowIdx, 1, row.sku,    FILL.skuMissing, true,  10);
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'missing' as const };
        writeCell(ws, rowIdx, i + 2, c.ambra, FILL.missingRow, false, 10, true);
      });
      rowIdx++;

    } else if (row.rowStatus === 'added') {
      writeCell(ws, rowIdx, 0, 'AIFY ONLY', FILL.addedRow, true,  10, false, '064E3B');
      writeCell(ws, rowIdx, 1, row.sku,     FILL.skuAdded,  true, 10);
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'added' as const };
        writeCell(ws, rowIdx, i + 2, c.aify, FILL.addedRow, false, 10, true);
      });
      rowIdx++;

    } else {
      // MATCH
      writeCell(ws, rowIdx, 0, 'MATCH', FILL.matched, false, 10);
      writeCell(ws, rowIdx, 1, row.sku, FILL.matched, true,  10);
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        writeCell(ws, rowIdx, i + 2, c.ambra || c.aify, FILL.matched, false, 10, true);
      });
      rowIdx++;
    }
  }

  setRef(ws, Math.max(rowIdx - 1, 1), columns.length + 1);
  ws['!cols'] = [{ wch: 12 }, { wch: 22 }, ...columns.map(() => ({ wch: 24 }))];
  return ws;
}

// ─── Workbook builder ─────────────────────────────────────────────────────────
function buildWorkbook(result: CompareResult): XLSX.WorkBook {
  const workbook  = XLSX.utils.book_new();
  const diffRows  = result.rows.filter((r) => r.rowStatus !== 'match');
  const { columnChangeCounts } = result;

  XLSX.utils.book_append_sheet(workbook,
    buildSummarySheet(result.stats, columnChangeCounts), 'Summary');
  XLSX.utils.book_append_sheet(workbook,
    buildDataSheet(diffRows, result.columns, columnChangeCounts), 'Differences');
  XLSX.utils.book_append_sheet(workbook,
    buildDataSheet(result.rows, result.columns, columnChangeCounts), 'Full Comparison');

  return workbook;
}

// ─── Button component ─────────────────────────────────────────────────────────
export default function DownloadButton({ result }: DownloadButtonProps) {
  const handleDownload = () => {
    if (!result) return;
    // bookType:'xls' (BIFF8) is the only format where SheetJS community edition
    // writes cell fill and font styles.  xlsx requires SheetJS Pro for styles.
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
