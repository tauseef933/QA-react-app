import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Akeneo QA Tool',
  description: 'Compare Aify and Ambra spreadsheet data',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-slate-50 text-gray-900 antialiased">
        {children}
      </body>
    </html>
  );
}
