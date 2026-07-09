'use client';

// xlsx-js-style is a fork of SheetJS that actually writes cell fill/font styles.
// The community SheetJS (xlsx) package silently strips the `s` property on write.
import XLSXStyle from 'xlsx-js-style';
import type { CompareResult, DiffRow } from '@/lib/compare';

type DownloadButtonProps = {
  result: CompareResult | null;
};

// ─── Colour palette (ARGB hex, matches on-screen Tailwind classes) ─────────────
// xlsx-js-style uses 6-digit RGB (no alpha prefix needed for fgColor)
const C = {
  // Header
  headerGray:     'D1D5DB', // gray-300    — normal column header
  headerAmber:    'FEF3C7', // amber-100   — column with ≥1 change

  // Row type accents
  skuChanged:     'FCD34D', // amber-300
  skuMissing:     'FCA5A5', // red-300
  skuAdded:       '6EE7B7', // emerald-300

  // AIFY row
  aifyBadge:      'FED7AA', // orange-200  — "AIFY" source badge
  aifyRowBg:      'FFF7ED', // orange-50   — unchanged Aify cells
  aifyCellBad:    'FECACA', // red-200     — changed cell (wrong value)
  aifyTextBad:    '991B1B', // red-800

  // AMBRA ✓ row
  ambraBadge:     '6EE7B7', // emerald-300 — "AMBRA ✓" source badge
  ambraRowBg:     'F0FDF4', // green-50    — unchanged Ambra cells
  ambraCellGood:  'A7F3D0', // emerald-200 — changed cell (correct value)
  ambraTextGood:  '065F46', // emerald-900

  // Misc
  missingRow:     'FEE2E2', // red-100
  addedRow:       'D1FAE5', // emerald-100
  white:          'FFFFFF',

  // Text accents for other row types
  textMissing:    '7F1D1D',
  textAdded:      '064E3B',
  titleBlue:      '3B5BDB',
} as const;

// ─── Style builders ───────────────────────────────────────────────────────────
function fill(rgb: string) {
  return { fill: { patternType: 'solid' as const, fgColor: { rgb } } };
}

function font(opts: { bold?: boolean; sz?: number; color?: string }) {
  return {
    font: {
      bold: opts.bold ?? false,
      sz:   opts.sz   ?? 10,
      ...(opts.color ? { color: { rgb: opts.color } } : {}),
    },
  };
}

function align(wrap = true) {
  return { alignment: { wrapText: wrap, vertical: 'top' as const } };
}

function style(
  fillRgb?: string,
  bold?: boolean,
  sz?: number,
  fontColor?: string,
  wrap?: boolean,
): object {
  return {
    ...(fillRgb ? fill(fillRgb) : {}),
    ...font({ bold, sz: sz ?? 10, color: fontColor }),
    ...align(wrap ?? false),
  };
}

// ─── Cell writer ──────────────────────────────────────────────────────────────
function cell(
  value: string | number,
  fillRgb?: string,
  bold?: boolean,
  fontColor?: string,
  wrap?: boolean,
): object {
  const t = typeof value === 'number' ? 'n' : 's';
  return { t, v: value, s: style(fillRgb, bold, 10, fontColor, wrap) };
}

// ─── Summary sheet ────────────────────────────────────────────────────────────
function buildSummarySheet(
  stats: CompareResult['stats'],
  columnChangeCounts: CompareResult['columnChangeCounts'],
): object {
  const ws: Record<string, unknown> = {};
  let r = 0;

  function wc(row: number, col: number, v: string | number, s?: object) {
    const ref = XLSXStyle.utils.encode_cell({ r: row, c: col });
    ws[ref] = { t: typeof v === 'number' ? 'n' : 's', v, ...(s ? { s } : {}) };
  }

  // Title
  wc(r, 0, 'Akeneo QA Tool — Comparison Report',
    style(C.titleBlue, true, 14, C.white));
  wc(r, 1, '', style(C.titleBlue));
  wc(r, 2, '', style(C.titleBlue));
  r += 2;

  // Stats header
  wc(r, 0, 'STATISTICS', style(C.headerGray, true, 11));
  wc(r, 1, 'COUNT',      style(C.headerGray, true, 11));
  wc(r, 2, 'NOTE',       style(C.headerGray, true, 11));
  r++;

  const statRows: [string, number, string][] = [
    ['Total SKUs',                  stats.total,   'DBEAFE'],
    ['✔  Matched (identical)',      stats.matched, 'F3F4F6'],
    ['⚠  Changed (values differ)',  stats.changed, 'FEF3C7'],
    ['✗  Missing in Aify',          stats.missing, 'FEE2E2'],
    ['+  Aify Only (not in Ambra)', stats.added,   'D1FAE5'],
  ];
  const statNotes = [
    'All unique SKUs across both files',
    'All values match — no action needed',
    'Values differ between Aify and Ambra',
    'SKU exists in Ambra but not in Aify',
    'SKU in Aify but does not exist in Ambra',
  ];
  for (let i = 0; i < statRows.length; i++) {
    const [label, count, bg] = statRows[i];
    wc(r, 0, label,       style(bg, true,  10));
    wc(r, 1, count,       style(bg, false, 10));
    wc(r, 2, statNotes[i],style(bg, false, 10));
    r++;
  }
  r++;

  // Colour legend
  wc(r, 0, 'COLOUR LEGEND', style(C.headerGray, true, 11));
  wc(r, 1, 'MEANING',       style(C.headerGray, true, 11));
  wc(r, 2, 'ACTION',        style(C.headerGray, true, 11));
  r++;

  const legend: [string, string, string, string][] = [
    [C.aifyBadge,    'AIFY row (orange)',            'Your current Aify file value',      'Review — may need updating'],
    [C.aifyCellBad,  'Changed cell — AIFY (red)',    'Wrong value in your file',          'Replace with the green Ambra value'],
    [C.ambraBadge,   'AMBRA ✓ row (green)',           'Ambra source-of-truth value',       'This is the correct target value'],
    [C.ambraCellGood,'Changed cell — AMBRA (green)', 'Correct value from Ambra',          'Copy this into your Aify file'],
    [C.missingRow,   'Missing row (light red)',       'SKU in Ambra only',                 'Add this SKU to your Aify file'],
    [C.addedRow,     'Aify Only row (light green)',   'SKU in Aify only',                  'Verify if this SKU should exist'],
    [C.skuChanged,   'SKU cell (amber)',              'SKU has at least one change',        'Check Differences sheet'],
    [C.headerAmber,  'Column header (amber + count)', 'Column has ≥1 difference',          'Count shown in brackets e.g. finish_color (12)'],
  ];
  for (const [bg, label, meaning, action] of legend) {
    wc(r, 0, label,   style(bg, true,  10));
    wc(r, 1, meaning, style(bg, false, 10));
    wc(r, 2, action,  style(bg, false, 10));
    r++;
  }
  r++;

  // Top changed columns
  const changedCols = Object.entries(columnChangeCounts)
    .filter(([, n]) => n > 0)
    .sort(([, a], [, b]) => b - a);

  if (changedCols.length > 0) {
    const totalChanged = stats.changed + stats.missing + stats.added;
    wc(r, 0, `TOP CHANGED COLUMNS  (${changedCols.length} with differences)`,
      style(C.headerGray, true, 11));
    wc(r, 1, 'CHANGES',        style(C.headerGray, true, 11));
    wc(r, 2, '% OF PROBLEM ROWS', style(C.headerGray, true, 11));
    r++;
    for (const [col, count] of changedCols) {
      const pct = totalChanged > 0 ? `${((count / totalChanged) * 100).toFixed(1)}%` : '—';
      wc(r, 0, col,   style('FEF3C7', false, 10));
      wc(r, 1, count, style('FEF3C7', true,  10));
      wc(r, 2, pct,   style('FEF3C7', false, 10));
      r++;
    }
  }

  ws['!ref']  = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r, c: 2 } });
  ws['!cols'] = [{ wch: 42 }, { wch: 12 }, { wch: 52 }];
  return ws;
}

// ─── Data sheet ───────────────────────────────────────────────────────────────
function buildDataSheet(
  inputRows: DiffRow[],
  columns: string[],
  columnChangeCounts: CompareResult['columnChangeCounts'],
): object {
  const ws: Record<string, unknown> = {};

  function wc(r: number, c: number, v: string | number, s?: object) {
    const ref = XLSXStyle.utils.encode_cell({ r, c });
    ws[ref] = { t: typeof v === 'number' ? 'n' : 's', v, ...(s ? { s } : {}) };
  }

  // ── Header row ──────────────────────────────────────────────────────────────
  wc(0, 0, 'Source', style(C.headerGray, true, 10));
  wc(0, 1, 'SKU',    style(C.headerGray, true, 10));
  columns.forEach((col, i) => {
    const count = columnChangeCounts[col] ?? 0;
    const label = count > 0 ? `${col}  (${count})` : col;
    const bg    = count > 0 ? C.headerAmber : C.headerGray;
    wc(0, i + 2, label, style(bg, true, 10));
  });

  // ── Data rows ────────────────────────────────────────────────────────────────
  let ri = 1;

  for (const row of inputRows) {
    if (row.rowStatus === 'changed') {
      // AIFY row — wrong values
      wc(ri, 0, 'AIFY',   style(C.aifyBadge,  true,  10, C.aifyTextBad));
      wc(ri, 1, row.sku,  style(C.skuChanged, true,  10));
      columns.forEach((col, i) => {
        const c   = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        const bg  = c.status === 'changed' ? C.aifyCellBad : C.aifyRowBg;
        const fc  = c.status === 'changed' ? C.aifyTextBad : undefined;
        wc(ri, i + 2, c.aify, style(bg, c.status === 'changed', 10, fc, true));
      });
      ri++;

      // AMBRA ✓ row — correct values
      wc(ri, 0, 'AMBRA ✓', style(C.ambraBadge,   true,  10, C.ambraTextGood));
      wc(ri, 1, row.sku,   style(C.skuChanged,   true,  10));
      columns.forEach((col, i) => {
        const c   = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        const bg  = c.status === 'changed' ? C.ambraCellGood : C.ambraRowBg;
        const fc  = c.status === 'changed' ? C.ambraTextGood : undefined;
        wc(ri, i + 2, c.ambra, style(bg, c.status === 'changed', 10, fc, true));
      });
      ri++;

    } else if (row.rowStatus === 'missing') {
      wc(ri, 0, 'MISSING',  style(C.missingRow, true,  10, C.textMissing));
      wc(ri, 1, row.sku,    style(C.skuMissing, true,  10));
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'missing' as const };
        wc(ri, i + 2, c.ambra, style(C.missingRow, false, 10, undefined, true));
      });
      ri++;

    } else if (row.rowStatus === 'added') {
      wc(ri, 0, 'AIFY ONLY', style(C.addedRow,  true,  10, C.textAdded));
      wc(ri, 1, row.sku,     style(C.skuAdded,  true,  10));
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'added' as const };
        wc(ri, i + 2, c.aify, style(C.addedRow, false, 10, undefined, true));
      });
      ri++;

    } else {
      // MATCH
      wc(ri, 0, 'MATCH', style(C.white, false, 10));
      wc(ri, 1, row.sku, style(C.white, true,  10));
      columns.forEach((col, i) => {
        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
        wc(ri, i + 2, c.ambra || c.aify, style(C.white, false, 10, undefined, true));
      });
      ri++;
    }
  }

  ws['!ref']  = XLSXStyle.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(ri - 1, 1), c: columns.length + 1 } });
  ws['!cols'] = [{ wch: 12 }, { wch: 22 }, ...columns.map(() => ({ wch: 24 }))];
  return ws;
}

// ─── Workbook & download ──────────────────────────────────────────────────────
function buildWorkbook(result: CompareResult): ReturnType<typeof XLSXStyle.utils.book_new> {
  const wb       = XLSXStyle.utils.book_new();
  const diffRows = result.rows.filter((r) => r.rowStatus !== 'match');
  const { columnChangeCounts } = result;

  XLSXStyle.utils.book_append_sheet(
    wb,
    buildSummarySheet(result.stats, columnChangeCounts) as XLSXStyle.WorkSheet,
    'Summary',
  );
  XLSXStyle.utils.book_append_sheet(
    wb,
    buildDataSheet(diffRows, result.columns, columnChangeCounts) as XLSXStyle.WorkSheet,
    'Differences',
  );
  XLSXStyle.utils.book_append_sheet(
    wb,
    buildDataSheet(result.rows, result.columns, columnChangeCounts) as XLSXStyle.WorkSheet,
    'Full Comparison',
  );

  return wb;
}

export default function DownloadButton({ result }: DownloadButtonProps) {
  const handleDownload = () => {
    if (!result) return;
    XLSXStyle.writeFile(
      buildWorkbook(result),
      'akeneo-qa-report.xlsx',
      { bookType: 'xlsx' },
    );
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
      Download QA Report (.xlsx)
    </button>
  );
}
