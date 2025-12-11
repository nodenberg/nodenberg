import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'エクセル帳票生成システム',
  description: 'エクセルテンプレートからExcel/PDF帳票を生成',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="ja">
      <body>{children}</body>
    </html>
  )
}
