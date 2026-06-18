'use client';

type StatsBarProps = {
  stats: {
    total: number;
    matched: number;
    changed: number;
    added: number;
    missing: number;
  };
};

type Card = {
  key: keyof StatsBarProps['stats'];
  label: string;
  sublabel: string;
  cardClass: string;
  numClass: string;
  dotClass: string;
};

const cards: Card[] = [
  { key: 'total',   label: 'Total SKUs',      sublabel: 'across both files',     cardClass: 'bg-white border-gray-200',       numClass: 'text-gray-900',    dotClass: 'bg-brand' },
  { key: 'matched', label: 'Matched',          sublabel: 'identical in both',     cardClass: 'bg-gray-50 border-gray-200',     numClass: 'text-gray-700',    dotClass: 'bg-gray-400' },
  { key: 'changed', label: 'Changed',          sublabel: 'values differ',         cardClass: 'bg-amber-50 border-amber-200',   numClass: 'text-amber-900',   dotClass: 'bg-amber-400' },
  { key: 'missing', label: 'Missing in Aify',  sublabel: 'SKU not in Aify file',  cardClass: 'bg-red-50 border-red-200',       numClass: 'text-red-900',     dotClass: 'bg-red-400' },
  { key: 'added',   label: 'Aify Only',        sublabel: 'SKU not in Ambra',      cardClass: 'bg-emerald-50 border-emerald-200', numClass: 'text-emerald-900', dotClass: 'bg-emerald-400' },
];

export default function StatsBar({ stats }: StatsBarProps) {
  return (
    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-5">
      {cards.map((card) => (
        <div key={card.key} className={`rounded-xl border p-4 shadow-sm ${card.cardClass}`}>
          <div className="mb-2 flex items-center gap-2">
            <span className={`h-2.5 w-2.5 shrink-0 rounded-full ${card.dotClass}`} aria-hidden="true" />
            <span className="text-xs font-semibold uppercase tracking-wide text-gray-600">{card.label}</span>
          </div>
          <p className={`text-3xl font-bold tabular-nums ${card.numClass}`}>{stats[card.key].toLocaleString()}</p>
          <p className="mt-1 text-xs text-gray-400">{card.sublabel}</p>
        </div>
      ))}
    </div>
  );
}
