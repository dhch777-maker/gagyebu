import type { Metadata } from 'next'
import './globals.css'

export const metadata: Metadata = {
  title: 'BUNT Kunstschule · 마곡원',
  description: '분트쿤스트슐레 마곡원 - 아동 미술 학원',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  )
}
