import type { Metadata } from 'next';
import { Inter } from 'next/font/google';
import './globals.css';

const inter = Inter({ subsets: ['latin'] });

export const metadata: Metadata = {
  title: 'ТД 1С - Excel',
  description: 'Transform your Excel files from 1C format',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ru">
      <body
        className={inter.className}
        style={{ backgroundColor: '#0f172a', minHeight: '100vh' }}
      >
        {children}
      </body>
    </html>
  );
}
