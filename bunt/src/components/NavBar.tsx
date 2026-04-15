'use client'

import { useRouter } from 'next/navigation'
import { createClient } from '@/lib/supabase'

interface Props { isAdmin?: boolean }

export default function NavBar({ isAdmin }: Props) {
  const router = useRouter()

  async function handleLogout() {
    const supabase = createClient()
    await supabase.auth.signOut()
    router.push('/')
    router.refresh()
  }

  return (
    <nav style={{
      background: '#0a0a0a',
      borderBottom: '1px solid #1e1e1e',
      padding: '12px 24px',
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      fontFamily: 'Georgia, serif',
    }}>
      <div style={{ fontSize: 14, letterSpacing: 5, color: '#d4a853', textTransform: 'uppercase' }}>
        BUNT
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 20 }}>
        {isAdmin && (
          <span style={{ fontSize: 9, letterSpacing: 2, color: '#d4a85360', textTransform: 'uppercase' }}>
            관리자 모드
          </span>
        )}
        <button onClick={handleLogout} style={{
          background: 'none', border: 'none',
          fontSize: 9, letterSpacing: 2, color: '#333',
          textTransform: 'uppercase', cursor: 'pointer',
          fontFamily: 'Georgia, serif',
        }}>로그아웃</button>
      </div>
    </nav>
  )
}
