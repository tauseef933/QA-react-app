'use client';

import { useCallback, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

type FileUploadProps = {
  label: string;
  onLoad: (rows: Record<string, string>[], columns: string[], fileName: string) => void;
  loaded: boolean;
  rowCount: number;
  fileName: string;
};

type SheetRow = (string | number | boolean | null | undefined)[];

function parseWorkbook(buffer: ArrayBuffer): {
  rows: Record<string, string>[];
  columns: string[];
} {
  const workbook = XLSX.read(buffer, { type: 'array', raw: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json<SheetRow>(sheet, { header: 1 });

  if (data.length === 0) {
    return { rows: [], columns: [] };
  }

  const columns = data[0].map((header) => String(header ?? '').trim());
  const rows: Record<string, string>[] = [];

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const record: Record<string, string> = {};
    for (let j = 0; j < columns.length; j++) {
      const value = rowData[j];
      record[columns[j]] = value != null ? String(value) : '';
    }
    rows.push(record);
  }

  return { rows, columns };
}

export default function FileUpload({
  label,
  onLoad,
  loaded,
  rowCount,
  fileName,
}: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);
  const [selectedFileName, setSelectedFileName] = useState('');

  const displayFileName = fileName || selectedFileName;

  const processFile = useCallback(
    async (file: File) => {
      if (!file.name.toLowerCase().endsWith('.xlsx')) return;

      setIsParsing(true);
      setSelectedFileName(file.name);

      try {
        const buffer = await file.arrayBuffer();
        const { rows, columns } = parseWorkbook(buffer);
        onLoad(rows, columns, file.name);
      } finally {
        setIsParsing(false);
      }
    },
    [onLoad],
  );

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) void processFile(file);
    event.target.value = '';
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragOver(false);
    const file = event.dataTransfer.files[0];
    if (file) void processFile(file);
  };

  const openFilePicker = () => inputRef.current?.click();

  return (
    <div className="flex flex-col gap-2">
      <span className="text-sm font-medium text-gray-700">{label}</span>

      <div
        role="button"
        tabIndex={0}
        onClick={loaded && !isParsing ? undefined : openFilePicker}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            openFilePicker();
          }
        }}
        onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
        onDragLeave={() => setIsDragOver(false)}
        onDrop={handleDrop}
        className={[
          'rounded-xl border-2 border-dashed p-6 text-center transition-colors',
          isDragOver
            ? 'border-brand bg-brand/5'
            : 'border-gray-300 hover:border-brand hover:bg-brand/5',
          loaded && !isParsing ? 'cursor-default' : 'cursor-pointer',
        ].join(' ')}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          onChange={handleFileChange}
          className="hidden"
        />

        {isParsing ? (
          <div className="flex flex-col items-center gap-3 py-2">
            <div className="h-8 w-8 animate-spin rounded-full border-2 border-gray-200 border-t-brand" aria-hidden="true" />
            <p className="text-sm text-gray-600">Parsing spreadsheet…</p>
          </div>
        ) : loaded ? (
          <div className="flex flex-col items-center gap-2 py-2">
            <div className="flex h-10 w-10 items-center justify-center rounded-full bg-emerald-100">
              <svg className="h-5 w-5 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5} aria-hidden="true">
                <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
              </svg>
            </div>
            <p className="text-sm font-medium text-gray-900">{displayFileName}</p>
            <p className="text-sm text-gray-500">{rowCount.toLocaleString()} row{rowCount === 1 ? '' : 's'}</p>
            <button
              type="button"
              onClick={(e) => { e.stopPropagation(); openFilePicker(); }}
              className="text-sm font-medium text-brand hover:underline"
            >
              Change file
            </button>
          </div>
        ) : (
          <div className="flex flex-col items-center gap-2 py-2">
            <div className={['flex h-10 w-10 items-center justify-center rounded-full', isDragOver ? 'bg-brand/10 text-brand' : 'bg-gray-100 text-gray-400'].join(' ')}>
              <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden="true">
                <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            </div>
            <p className="text-sm font-medium text-gray-900">Drop .xlsx file here or click to browse</p>
            <p className="text-xs text-gray-500">Excel (.xlsx) only</p>
          </div>
        )}
      </div>
    </div>
  );
}
