export type CellStatus = 'match' | 'changed' | 'added' | 'missing';

export type DiffCell = {
  aify: string;
  ambra: string;
  status: CellStatus;
};

export type DiffRow = {
  sku: string;
  rowStatus: CellStatus;
  cells: Record<string, DiffCell>;
};

export type CompareResult = {
  columns: string[];
  rows: DiffRow[];
  /** Number of 'changed' cells per column across all rows. */
  columnChangeCounts: Record<string, number>;
  stats: {
    total: number;
    matched: number;
    changed: number;
    added: number;
    missing: number;
  };
};

function normalizeValue(value: string | undefined): string {
  return (value ?? '').toString().trim();
}

function sortBySku(
  rows: Record<string, string>[],
  skuKey: string,
): Record<string, string>[] {
  return [...rows].sort((a, b) =>
    normalizeValue(a[skuKey]).localeCompare(normalizeValue(b[skuKey])),
  );
}

/**
 * Find the actual column key that represents SKU (case-insensitive).
 * Falls back to 'sku' if none is found.
 */
export function detectSkuColumn(columns: string[]): string {
  return columns.find((c) => c.toLowerCase() === 'sku') ?? 'sku';
}

export function compareFiles(
  aifyRows: Record<string, string>[],
  ambraRows: Record<string, string>[],
  columns: string[],
  skuColumnName: string,
): CompareResult {
  // Sort columns alphabetically so display order is consistent regardless of
  // which order the columns appear in either file.
  const sortedColumns = [...columns].sort((a, b) => a.localeCompare(b));

  // Build SKU-keyed maps from both files (sorted before indexing)
  const aifyMap = new Map<string, Record<string, string>>();
  for (const row of sortBySku(aifyRows, skuColumnName)) {
    const key = normalizeValue(row[skuColumnName]);
    if (key) aifyMap.set(key, row);
  }

  const ambraMap = new Map<string, Record<string, string>>();
  for (const row of sortBySku(ambraRows, skuColumnName)) {
    const key = normalizeValue(row[skuColumnName]);
    if (key) ambraMap.set(key, row);
  }

  // Union of all SKUs, sorted
  const allSkus = new Set([
    ...Array.from(aifyMap.keys()),
    ...Array.from(ambraMap.keys()),
  ]);
  const sortedSkus = Array.from(allSkus).sort((a, b) => a.localeCompare(b));

  const rows: DiffRow[] = [];
  const columnChangeCounts: Record<string, number> = {};
  for (const col of sortedColumns) {
    columnChangeCounts[col] = 0;
  }

  let matched = 0;
  let changed = 0;
  let added = 0;
  let missing = 0;

  for (const sku of sortedSkus) {
    const aifyRow = aifyMap.get(sku);
    const ambraRow = ambraMap.get(sku);

    if (aifyRow && ambraRow) {
      // SKU present in both — compare cell by cell
      const cells: Record<string, DiffCell> = {};
      let rowStatus: CellStatus = 'match';

      for (const col of sortedColumns) {
        const aifyVal = normalizeValue(aifyRow[col]);
        const ambraVal = normalizeValue(ambraRow[col]);
        const status: CellStatus = aifyVal === ambraVal ? 'match' : 'changed';

        if (status === 'changed') {
          rowStatus = 'changed';
          columnChangeCounts[col]++;
        }

        cells[col] = { aify: aifyVal, ambra: ambraVal, status };
      }

      if (rowStatus === 'match') {
        matched++;
      } else {
        changed++;
      }

      rows.push({ sku, rowStatus, cells });
    } else if (aifyRow) {
      // SKU only in Aify — extra row not in Ambra
      const cells: Record<string, DiffCell> = {};
      for (const col of sortedColumns) {
        cells[col] = { aify: normalizeValue(aifyRow[col]), ambra: '', status: 'added' };
      }
      added++;
      rows.push({ sku, rowStatus: 'added', cells });
    } else if (ambraRow) {
      // SKU only in Ambra — missing from Aify
      const cells: Record<string, DiffCell> = {};
      for (const col of sortedColumns) {
        cells[col] = { aify: '', ambra: normalizeValue(ambraRow[col]), status: 'missing' };
      }
      missing++;
      rows.push({ sku, rowStatus: 'missing', cells });
    }
  }

  return {
    columns: sortedColumns,
    rows,
    columnChangeCounts,
    stats: { total: allSkus.size, matched, changed, added, missing },
  };
}
