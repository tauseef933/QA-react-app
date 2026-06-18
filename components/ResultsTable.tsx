'use client';

import { Fragment, useMemo, useState } from 'react';
import type { CellStatus, CompareResult } from '@/lib/compare';

type ResultsTableProps = {
  result: CompareResult;
  visibleColumns: string[];
};

type FilterTab = 'all' | CellStatus;

const ROWS_PER_PAGE = 50;

const filterTabs: { key: FilterTab; label: string }[] = [
  { key: 'all',     label: 'All' },
  { key: 'changed', label: 'Changed' },
  { key: 'missing', label: 'Missing' },
  { key: 'added',   label: 'Aify Only' },
  { key: 'match',   label: 'Matched' },
];

function Empty() {
  return <em className="select-none text-gray-300 not-italic">—</em>;
}

/**
 * Clips cell text to 3 lines so tall descriptions don't blow out row height.
 * Users can still read the full value in the Excel download.
 */
function Cell({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <div className={['line-clamp-3 break-words text-xs leading-relaxed', className].filter(Boolean).join(' ')}>
      {children}
    </div>
  );
}

export default function ResultsTable({ result, visibleColumns }: ResultsTableProps) {
  const [filter, setFilter] = useState<FilterTab>('all');
  const [page, setPage] = useState(1);

  const counts = useMemo(() => ({
    all:     result.rows.length,
    changed: result.rows.filter((r) => r.rowStatus === 'changed').length,
    missing: result.rows.filter((r) => r.rowStatus === 'missing').length,
    added:   result.rows.filter((r) => r.rowStatus === 'added').length,
    match:   result.rows.filter((r) => r.rowStatus === 'match').length,
  }), [result.rows]);

  const filteredRows = useMemo(() => {
    if (filter === 'all') return result.rows;
    return result.rows.filter((row) => row.rowStatus === filter);
  }, [result.rows, filter]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / ROWS_PER_PAGE));
  const currentPage = Math.min(page, totalPages);

  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * ROWS_PER_PAGE;
    return filteredRows.slice(start, start + ROWS_PER_PAGE);
  }, [filteredRows, currentPage]);

  const handleFilterChange = (next: FilterTab) => { setFilter(next); setPage(1); };

  return (
    <div className="flex flex-col gap-4">

      {/* Filter tabs */}
      <div className="flex flex-wrap gap-2">
        {filterTabs.map((tab) => {
          const isActive = filter === tab.key;
          return (
            <button
              key={tab.key}
              type="button"
              onClick={() => handleFilterChange(tab.key)}
              className={['inline-flex items-center gap-2 rounded-full px-4 py-1.5 text-sm font-medium transition-all',
                isActive ? 'bg-brand text-white shadow-sm' : 'border border-gray-200 bg-white text-gray-700 hover:border-brand/40 hover:bg-blue-50',
              ].join(' ')}
            >
              {tab.label}
              <span className={['rounded-full px-1.5 py-0.5 text-xs font-semibold tabular-nums',
                isActive ? 'bg-white/25 text-white' : 'bg-gray-100 text-gray-600',
              ].join(' ')}>
                {counts[tab.key]}
              </span>
            </button>
          );
        })}
      </div>

      {/* Legend */}
      <div className="flex flex-wrap items-center gap-3 rounded-lg border border-gray-100 bg-gray-50 px-4 py-2.5 text-xs text-gray-600">
        <span className="font-semibold uppercase tracking-wide text-gray-500">How to read:</span>
        <span className="flex items-center gap-1.5"><span className="inline-block rounded bg-orange-200 px-2 py-0.5 text-[10px] font-bold text-orange-900">Aify</span> your file row</span>
        <span className="flex items-center gap-1.5"><span className="inline-block rounded bg-emerald-300 px-2 py-0.5 text-[10px] font-bold text-emerald-900">Ambra ✓</span> source of truth row</span>
        <span className="flex items-center gap-1.5"><span className="inline-block h-3 w-10 rounded border border-red-200 bg-red-100" /> old value (Aify)</span>
        <span className="flex items-center gap-1.5"><span className="inline-block h-3 w-10 rounded border border-emerald-300 bg-emerald-100" /> correct value (Ambra)</span>
        <span className="ml-auto text-gray-400">Drag ↕ bottom-right corner to resize table</span>
      </div>

      <p className="text-xs text-gray-500">
        Showing <span className="font-semibold text-gray-700">{paginatedRows.length.toLocaleString()}</span> of{' '}
        <span className="font-semibold text-gray-700">{filteredRows.length.toLocaleString()}</span> SKUs
        <span className="ml-2 text-gray-400">(cell text clipped to 3 lines — full values in downloaded report)</span>
      </p>

      {/*
        Table container:
        - h-[480px]   → shows ~12 rows by default
        - min-h-[180px] → prevents collapsing too small
        - resize-y    → user can drag the bottom edge to make it taller/shorter
        - overflow-auto → enables independent horizontal + vertical scrolling
      */}
      <div className="h-[480px] min-h-[180px] resize-y overflow-auto rounded-xl border border-gray-200 shadow-sm">
        <table className="min-w-full border-collapse text-sm">
          <thead>
            <tr className="border-b-2 border-gray-200 bg-gray-50">
              {/* Both top-0 AND left-0: stays fixed while scrolling either direction */}
              <th className="sticky left-0 top-0 z-30 w-[80px] shrink-0 border-r border-gray-200 bg-gray-50 px-3 py-2.5 text-left text-xs font-semibold uppercase tracking-wider text-gray-400">
                Source
              </th>
              <th className="sticky left-[80px] top-0 z-30 min-w-[150px] border-r border-gray-200 bg-gray-50 px-4 py-2.5 text-left text-xs font-semibold uppercase tracking-wider text-gray-500">
                SKU
              </th>
              {visibleColumns.map((column) => {
                const changeCount = result.columnChangeCounts[column] ?? 0;
                const hasChanges = changeCount > 0;
                return (
                  <th
                    key={column}
                    title={`${column}${hasChanges ? ` — ${changeCount} change${changeCount === 1 ? '' : 's'}` : ''}`}
                    className={[
                      'sticky top-0 z-10 min-w-[150px] max-w-[200px] px-3 py-2.5 text-left text-xs font-semibold uppercase tracking-wider',
                      hasChanges ? 'bg-amber-50 text-amber-800' : 'bg-gray-50 text-gray-500',
                    ].join(' ')}
                  >
                    <div className="flex items-center gap-1.5">
                      <span className="truncate">{column}</span>
                      {hasChanges && (
                        <span className="shrink-0 rounded-full bg-amber-300 px-1.5 py-0.5 text-[10px] font-bold tabular-nums text-amber-900">
                          {changeCount}
                        </span>
                      )}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>

          <tbody>
            {paginatedRows.length === 0 ? (
              <tr>
                <td colSpan={visibleColumns.length + 2} className="px-4 py-14 text-center">
                  <div className="flex flex-col items-center gap-3 text-gray-400">
                    <svg className="h-10 w-10 text-gray-200" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                    </svg>
                    <span className="text-sm">No rows match this filter</span>
                  </div>
                </td>
              </tr>
            ) : (
              paginatedRows.map((row) => {

                /* ── CHANGED: two rows — Aify (old) then Ambra (correct) ────── */
                if (row.rowStatus === 'changed') {
                  return (
                    <Fragment key={row.sku}>
                      {/* Aify row */}
                      <tr className="border-b border-gray-100 bg-white">
                        <td className="sticky left-0 z-10 w-[80px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                          <span className="inline-block rounded bg-orange-200 px-1.5 py-0.5 text-[10px] font-bold uppercase text-orange-900">Aify</span>
                        </td>
                        <td className="sticky left-[80px] z-10 min-w-[150px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                          <span className="break-all text-xs font-bold text-gray-900">{row.sku}</span>
                        </td>
                        {visibleColumns.map((col) => {
                          const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
                          return (
                            <td
                              key={`${row.sku}__aify__${col}`}
                              className={['min-w-[150px] max-w-[200px] px-3 py-2 align-top',
                                c.status === 'changed' ? 'bg-red-100 ring-1 ring-inset ring-red-200' : '',
                              ].join(' ')}
                            >
                              {c.status === 'changed'
                                ? <Cell className="font-medium text-red-800">{c.aify || <Empty />}</Cell>
                                : <Cell className="text-gray-400">{c.aify || <Empty />}</Cell>}
                            </td>
                          );
                        })}
                      </tr>

                      {/* Ambra row */}
                      <tr className="border-b-2 border-gray-200 bg-white">
                        <td className="sticky left-0 z-10 w-[80px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                          <span className="inline-block rounded bg-emerald-300 px-1.5 py-0.5 text-[10px] font-bold uppercase text-emerald-900">Ambra ✓</span>
                        </td>
                        <td className="sticky left-[80px] z-10 min-w-[150px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                          <span className="break-all text-xs font-bold text-gray-900">{row.sku}</span>
                        </td>
                        {visibleColumns.map((col) => {
                          const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
                          return (
                            <td
                              key={`${row.sku}__ambra__${col}`}
                              className={['min-w-[150px] max-w-[200px] px-3 py-2 align-top',
                                c.status === 'changed' ? 'bg-emerald-100 ring-1 ring-inset ring-emerald-300' : '',
                              ].join(' ')}
                            >
                              {c.status === 'changed'
                                ? <Cell className="font-semibold text-emerald-900">{c.ambra || <Empty />}</Cell>
                                : <Cell className="text-gray-500">{c.ambra || <Empty />}</Cell>}
                            </td>
                          );
                        })}
                      </tr>
                    </Fragment>
                  );
                }

                /* ── MISSING ──────────────────────────────────────────────── */
                if (row.rowStatus === 'missing') {
                  return (
                    <tr key={row.sku} className="border-b border-gray-100 bg-white">
                      <td className="sticky left-0 z-10 w-[80px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                        <span className="inline-block rounded bg-red-200 px-1.5 py-0.5 text-[10px] font-bold uppercase text-red-900">Missing</span>
                      </td>
                      <td className="sticky left-[80px] z-10 min-w-[150px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                        <span className="break-all text-xs font-bold text-gray-900">{row.sku}</span>
                      </td>
                      {visibleColumns.map((col) => {
                        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'missing' as const };
                        return (
                          <td key={`${row.sku}__${col}`} className="min-w-[150px] max-w-[200px] px-3 py-2 align-top">
                            <Cell className="text-gray-700">{c.ambra || <Empty />}</Cell>
                          </td>
                        );
                      })}
                    </tr>
                  );
                }

                /* ── ADDED (Aify only) ────────────────────────────────────── */
                if (row.rowStatus === 'added') {
                  return (
                    <tr key={row.sku} className="border-b border-gray-100 bg-white">
                      <td className="sticky left-0 z-10 w-[80px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                        <span className="inline-block rounded bg-emerald-200 px-1.5 py-0.5 text-[10px] font-bold uppercase text-emerald-900">Aify Only</span>
                      </td>
                      <td className="sticky left-[80px] z-10 min-w-[150px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                        <span className="break-all text-xs font-bold text-gray-900">{row.sku}</span>
                      </td>
                      {visibleColumns.map((col) => {
                        const c = row.cells[col] ?? { aify: '', ambra: '', status: 'added' as const };
                        return (
                          <td key={`${row.sku}__${col}`} className="min-w-[150px] max-w-[200px] px-3 py-2 align-top">
                            <Cell className="text-gray-700">{c.aify || <Empty />}</Cell>
                          </td>
                        );
                      })}
                    </tr>
                  );
                }

                /* ── MATCH ────────────────────────────────────────────────── */
                return (
                  <tr key={row.sku} className="border-b border-gray-100 bg-white hover:bg-gray-50">
                    <td className="sticky left-0 z-10 w-[80px] border-r border-gray-100 bg-white px-3 py-2 align-top">
                      <span className="inline-block rounded bg-gray-100 px-1.5 py-0.5 text-[10px] font-medium uppercase text-gray-500">Match</span>
                    </td>
                    <td className="sticky left-[80px] z-10 min-w-[150px] border-r border-gray-200 bg-white px-3 py-2 align-top">
                      <span className="break-all text-xs font-semibold text-gray-800">{row.sku}</span>
                    </td>
                    {visibleColumns.map((col) => {
                      const c = row.cells[col] ?? { aify: '', ambra: '', status: 'match' as const };
                      return (
                        <td key={`${row.sku}__${col}`} className="min-w-[150px] max-w-[200px] px-3 py-2 align-top">
                          <Cell className="text-gray-600">{c.ambra || c.aify || <Empty />}</Cell>
                        </td>
                      );
                    })}
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div className="flex items-center justify-between rounded-xl border border-gray-200 bg-white px-4 py-3 shadow-sm">
        <p className="text-sm text-gray-600">
          Page <span className="font-semibold">{currentPage}</span> of{' '}
          <span className="font-semibold">{totalPages}</span>
          <span className="ml-2 text-gray-400 tabular-nums">({filteredRows.length.toLocaleString()} SKUs)</span>
        </p>
        <div className="flex gap-1.5">
          {[
            { label: 'First',   action: () => setPage(1),                                  disabled: currentPage <= 1,          cls: 'px-2.5 text-xs' },
            { label: '← Prev',  action: () => setPage((p) => Math.max(1, p - 1)),          disabled: currentPage <= 1,          cls: 'px-3 text-sm' },
            { label: 'Next →',  action: () => setPage((p) => Math.min(totalPages, p + 1)), disabled: currentPage >= totalPages, cls: 'px-3 text-sm' },
            { label: 'Last',    action: () => setPage(totalPages),                          disabled: currentPage >= totalPages, cls: 'px-2.5 text-xs' },
          ].map((btn) => (
            <button key={btn.label} type="button" onClick={btn.action} disabled={btn.disabled}
              className={`rounded-md border border-gray-200 py-1.5 font-medium text-gray-700 transition-colors hover:bg-gray-50 disabled:cursor-not-allowed disabled:opacity-40 ${btn.cls}`}>
              {btn.label}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
