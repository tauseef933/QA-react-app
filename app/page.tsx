'use client';

import { useMemo, useState } from 'react';
import DownloadButton from '@/components/DownloadButton';
import FileUpload from '@/components/FileUpload';
import ResultsTable from '@/components/ResultsTable';
import StatsBar from '@/components/StatsBar';
import { compareFiles, detectSkuColumn, type CompareResult } from '@/lib/compare';

export default function Home() {
  const [aifyRows, setAifyRows] = useState<Record<string, string>[]>([]);
  const [aifyFileName, setAifyFileName] = useState('');
  const [aifyRowCount, setAifyRowCount] = useState(0);

  const [aifyColumns, setAifyColumns] = useState<string[]>([]);

  const [ambraRows, setAmbraRows] = useState<Record<string, string>[]>([]);
  const [ambraColumns, setAmbraColumns] = useState<string[]>([]);
  const [ambraFileName, setAmbraFileName] = useState('');
  const [ambraRowCount, setAmbraRowCount] = useState(0);

  const [result, setResult] = useState<CompareResult | null>(null);
  const [isComparing, setIsComparing] = useState(false);
  const [visibleColumns, setVisibleColumns] = useState<string[]>([]);
  const [columnSearch, setColumnSearch] = useState('');
  const [changesOnly, setChangesOnly] = useState(false);

  const bothFilesLoaded = aifyRowCount > 0 && ambraRowCount > 0;

  const skuColumnName = useMemo(
    () => detectSkuColumn(ambraColumns.length ? ambraColumns : aifyColumns),
    [ambraColumns, aifyColumns],
  );

  const comparisonColumns = useMemo(() => {
    const skuLower = skuColumnName.toLowerCase();

    // Start with Ambra columns (source of truth drives the order)
    const seen = new Map<string, string>(); // lowerKey → display name
    for (const c of ambraColumns) {
      if (c && c.toLowerCase() !== skuLower) seen.set(c.toLowerCase(), c);
    }

    // Append any Aify-only columns that don't already exist in Ambra.
    // This ensures columns like "collection" present only in Aify still appear.
    for (const c of aifyColumns) {
      if (c && c.toLowerCase() !== skuLower && !seen.has(c.toLowerCase())) {
        seen.set(c.toLowerCase(), c);
      }
    }

    return Array.from(seen.values());
  }, [ambraColumns, aifyColumns, skuColumnName]);

  const displayedColumns = useMemo(() => {
    let cols = visibleColumns;
    if (changesOnly && result) {
      cols = cols.filter((col) => (result.columnChangeCounts[col] ?? 0) > 0);
    }
    const search = columnSearch.trim().toLowerCase();
    if (search) cols = cols.filter((col) => col.toLowerCase().includes(search));
    return cols;
  }, [visibleColumns, changesOnly, result, columnSearch]);

  const handleAifyLoad = (rows: Record<string, string>[], cols: string[], fileName: string) => {
    setAifyRows(rows);
    setAifyColumns(cols);
    setAifyFileName(fileName);
    setAifyRowCount(rows.length);
    setResult(null);
  };

  const handleAmbraLoad = (rows: Record<string, string>[], columns: string[], fileName: string) => {
    setAmbraRows(rows);
    setAmbraColumns(columns);
    setAmbraFileName(fileName);
    setAmbraRowCount(rows.length);
    setResult(null);
  };

  const handleCompare = async () => {
    setIsComparing(true);
    await new Promise<void>((resolve) => { window.setTimeout(resolve, 0); });
    const r = compareFiles(aifyRows, ambraRows, comparisonColumns, skuColumnName);
    setResult(r);
    setVisibleColumns(r.columns);
    setChangesOnly(false);
    setColumnSearch('');
    setIsComparing(false);
  };

  const handleShowAll = () => {
    if (result) setVisibleColumns(result.columns);
    setChangesOnly(false);
    setColumnSearch('');
  };

  return (
    <div className="min-h-screen bg-slate-50">

      {/* Header */}
      <header className="bg-brand text-white shadow-lg">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg bg-white/20">
              <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden="true">
                <path strokeLinecap="round" strokeLinejoin="round" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-bold leading-tight tracking-tight">Akeneo QA Tool</h1>
              <p className="text-xs leading-tight text-white/70">Product Data Quality Check</p>
            </div>
          </div>
          <div className="hidden items-center gap-2 rounded-full bg-white/10 px-3 py-1.5 sm:flex">
            <span className="text-sm font-medium text-white/90">aify</span>
            <span className="text-white/40">×</span>
            <span className="text-sm font-medium text-white/90">Ambra</span>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-7xl space-y-8 px-4 py-8 sm:px-6">

        {/* Step 1: Upload */}
        <section>
          <SectionLabel step={1} text="Upload Files" />
          <div className="grid gap-4 md:grid-cols-2">
            <FileUpload label="Aify File (Your Template)"       onLoad={handleAifyLoad}  loaded={aifyRowCount > 0}  rowCount={aifyRowCount}  fileName={aifyFileName} />
            <FileUpload label="Ambra File (Source of Truth)"    onLoad={handleAmbraLoad} loaded={ambraRowCount > 0} rowCount={ambraRowCount} fileName={ambraFileName} />
          </div>
        </section>

        {/* Step 2: Compare */}
        {bothFilesLoaded && (
          <section>
            <SectionLabel step={2} text="Run Comparison" />
            <div className="flex flex-wrap items-center gap-4 rounded-xl border border-gray-200 bg-white p-4 shadow-sm">
              <div className="min-w-0 flex-1 text-sm text-gray-600">
                <p>
                  Comparing <span className="font-semibold text-gray-900">{aifyFileName}</span> against <span className="font-semibold text-gray-900">{ambraFileName}</span>
                  {skuColumnName !== 'sku' && (
                    <span className="ml-2 text-gray-400">· SKU column: <code className="rounded bg-gray-100 px-1 text-xs text-gray-700">{skuColumnName}</code></span>
                  )}
                </p>
                <p className="mt-0.5 text-xs text-gray-400">
                  Aify: <span className="font-semibold text-gray-600">{aifyColumns.length}</span> columns ·{' '}
                  Ambra: <span className="font-semibold text-gray-600">{ambraColumns.length}</span> columns ·{' '}
                  Comparing: <span className="font-semibold text-gray-600">{comparisonColumns.length}</span> columns
                </p>
              </div>
              <button
                type="button"
                onClick={() => void handleCompare()}
                disabled={isComparing}
                className="flex shrink-0 items-center gap-2 rounded-lg bg-brand px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition-all hover:opacity-90 disabled:cursor-not-allowed disabled:bg-gray-300 disabled:shadow-none"
              >
                {isComparing ? (
                  <><div className="h-4 w-4 animate-spin rounded-full border-2 border-white/30 border-t-white" />Comparing…</>
                ) : (
                  <><svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5} aria-hidden="true"><path strokeLinecap="round" strokeLinejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" /></svg>Compare Files</>
                )}
              </button>
            </div>
          </section>
        )}

        {/* Step 3: Results */}
        {result && (
          <section className="space-y-6">
            <SectionLabel step={3} text="QA Report" />

            <StatsBar stats={result.stats} />

            {/* Column filter bar */}
            <div className="flex flex-wrap items-center gap-3 rounded-xl border border-gray-200 bg-white px-4 py-3 shadow-sm">
              <div className="flex min-w-[200px] flex-1 items-center gap-2 rounded-lg border border-gray-200 bg-gray-50 px-3 py-1.5">
                <svg className="h-4 w-4 shrink-0 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden="true">
                  <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                </svg>
                <input
                  type="text"
                  value={columnSearch}
                  onChange={(e) => setColumnSearch(e.target.value)}
                  placeholder="Search columns…"
                  className="flex-1 bg-transparent text-sm text-gray-900 outline-none placeholder:text-gray-400"
                />
              </div>
              <button type="button" onClick={handleShowAll} className="rounded-lg border border-gray-200 px-3 py-1.5 text-sm font-medium text-gray-700 transition-colors hover:bg-gray-50">
                Show All
              </button>
              <button
                type="button"
                onClick={() => setChangesOnly((p) => !p)}
                className={['rounded-lg border px-3 py-1.5 text-sm font-medium transition-colors',
                  changesOnly ? 'border-amber-300 bg-amber-50 text-amber-800' : 'border-gray-200 text-gray-700 hover:bg-gray-50',
                ].join(' ')}
              >
                {changesOnly ? '✓ Changes Only' : 'Changes Only'}
              </button>
              <span className="ml-auto text-xs tabular-nums text-gray-400">
                {displayedColumns.length} / {result.columns.length} columns
              </span>
            </div>

            <ResultsTable result={result} visibleColumns={displayedColumns} />

            <div className="flex flex-wrap items-center justify-between gap-4 rounded-xl border border-gray-200 bg-white p-4 shadow-sm">
              <div>
                <p className="text-sm font-semibold text-gray-900">Download Full QA Report</p>
                <p className="mt-0.5 text-xs text-gray-500">Excel file with Summary, Differences, and Full Comparison sheets</p>
              </div>
              <DownloadButton result={result} />
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

function SectionLabel({ step, text }: { step: number; text: string }) {
  return (
    <div className="mb-4 flex items-center gap-2">
      <span className="flex h-6 w-6 shrink-0 items-center justify-center rounded-full bg-brand text-xs font-bold text-white">{step}</span>
      <h2 className="text-sm font-semibold uppercase tracking-wide text-gray-500">{text}</h2>
    </div>
  );
}
