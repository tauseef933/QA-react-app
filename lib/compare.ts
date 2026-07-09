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

/**
 * Cleans a value for display — trims, collapses whitespace, and replaces
 * invisible Unicode space variants (non-breaking space, BOM, zero-width).
 */
function normalizeDisplay(value: string | undefined): string {
  if (value == null) return '';
  return value
    .toString()
    .replace(/[\u00A0\u00AD\u200B\uFEFF]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normalizes a value for semantic comparison:
 *   - Case-insensitive  ("Brass" === "brass")
 *   - Numeric coercion  ("10.50" === "10.5", "1,000" === "1000")
 *   - Boolean flattening ("Yes" / "1" / "TRUE" all === "true")
 *   - Whitespace normalization
 *
 * The DISPLAY value is kept separate so users still see the original text.
 */
function normalizeForComparison(raw: string | undefined): string {
  const v = normalizeDisplay(raw);
  if (!v) return '';

  // Numeric normalization — strip thousands-separator commas first
  const stripped = v.replace(/,/g, '');
  const num = Number(stripped);
  if (stripped !== '' && !isNaN(num) && isFinite(num)) {
    return String(parseFloat(stripped));
  }

  // Boolean normalization
  const lower = v.toLowerCase();
  if (lower === 'true' || lower === 'yes' || lower === '1') return 'true';
  if (lower === 'false' || lower === 'no' || lower === '0') return 'false';

  // General case-insensitive comparison
  return lower;
}

/**
 * Returns a lookup function for a single row that first tries an exact key
 * match and falls back to a case-insensitive match.  This lets Aify files
 * with slightly different column-name casing (e.g. "Categories" vs "categories")
 * still be matched against Ambra column names.
 */
function buildRowLookup(row: Record<string, string>): (col: string) => string {
  const lowerMap = new Map<string, string>();
  for (const [k, v] of Object.entries(row)) {
    // Only insert the first match for each lower-cased key so the original
    // exact-match path still wins (direct property lookup below).
    if (!lowerMap.has(k.toLowerCase())) {
      lowerMap.set(k.toLowerCase(), v);
    }
  }
  return (col: string): string => {
    if (col in row) return row[col];
    return lowerMap.get(col.toLowerCase()) ?? '';
  };
}

function sortBySku(
  rows: Record<string, string>[],
  skuKey: string,
): Record<string, string>[] {
  return [...rows].sort((a, b) => {
    const aVal = normalizeDisplay(buildRowLookup(a)(skuKey));
    const bVal = normalizeDisplay(buildRowLookup(b)(skuKey));
    return aVal.localeCompare(bVal);
  });
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
  // Deduplicate and sort columns alphabetically for consistent display order.
  const sortedColumns = Array.from(new Set(columns)).sort((a, b) => a.localeCompare(b));

  // Build SKU-keyed maps.  Keys are lowercased so SKU casing differences
  // between files don't prevent rows from matching.
  const aifyMap = new Map<string, Record<string, string>>();
  for (const row of sortBySku(aifyRows, skuColumnName)) {
    const key = normalizeDisplay(buildRowLookup(row)(skuColumnName)).toLowerCase();
    if (key) aifyMap.set(key, row);
  }

  const ambraMap = new Map<string, Record<string, string>>();
  for (const row of sortBySku(ambraRows, skuColumnName)) {
    const key = normalizeDisplay(buildRowLookup(row)(skuColumnName)).toLowerCase();
    if (key) ambraMap.set(key, row);
  }

  const allSkus = new Set(Array.from(aifyMap.keys()).concat(Array.from(ambraMap.keys())));
  const sortedSkus = Array.from(allSkus).sort((a, b) => a.localeCompare(b));

  const rows: DiffRow[] = [];
  const columnChangeCounts: Record<string, number> = {};
  for (const col of sortedColumns) columnChangeCounts[col] = 0;

  let matched = 0;
  let changed = 0;
  let added = 0;
  let missing = 0;

  for (const sku of sortedSkus) {
    const aifyRow = aifyMap.get(sku);
    const ambraRow = ambraMap.get(sku);

    if (aifyRow && ambraRow) {
      const aifyGet = buildRowLookup(aifyRow);
      const ambraGet = buildRowLookup(ambraRow);

      // Display SKU from Ambra (source of truth); fall back to Aify or key
      const displaySku =
        normalizeDisplay(ambraGet(skuColumnName)) ||
        normalizeDisplay(aifyGet(skuColumnName)) ||
        sku;

      const cells: Record<string, DiffCell> = {};
      let rowStatus: CellStatus = 'match';

      for (const col of sortedColumns) {
        const aifyRaw = aifyGet(col);
        const ambraRaw = ambraGet(col);
        const aifyDisplay = normalizeDisplay(aifyRaw);
        const ambraDisplay = normalizeDisplay(ambraRaw);
        const status: CellStatus =
          normalizeForComparison(aifyRaw) === normalizeForComparison(ambraRaw)
            ? 'match'
            : 'changed';

        if (status === 'changed') {
          rowStatus = 'changed';
          columnChangeCounts[col]++;
        }

        cells[col] = { aify: aifyDisplay, ambra: ambraDisplay, status };
      }

      if (rowStatus === 'match') matched++;
      else changed++;

      rows.push({ sku: displaySku, rowStatus, cells });
    } else if (aifyRow) {
      const aifyGet = buildRowLookup(aifyRow);
      const displaySku = normalizeDisplay(aifyGet(skuColumnName)) || sku;
      const cells: Record<string, DiffCell> = {};
      for (const col of sortedColumns) {
        cells[col] = { aify: normalizeDisplay(aifyGet(col)), ambra: '', status: 'added' };
      }
      added++;
      rows.push({ sku: displaySku, rowStatus: 'added', cells });
    } else if (ambraRow) {
      const ambraGet = buildRowLookup(ambraRow);
      const displaySku = normalizeDisplay(ambraGet(skuColumnName)) || sku;
      const cells: Record<string, DiffCell> = {};
      for (const col of sortedColumns) {
        cells[col] = { aify: '', ambra: normalizeDisplay(ambraGet(col)), status: 'missing' };
      }
      missing++;
      rows.push({ sku: displaySku, rowStatus: 'missing', cells });
    }
  }

  return {
    columns: sortedColumns,
    rows,
    columnChangeCounts,
    stats: { total: allSkus.size, matched, changed, added, missing },
  };
}
