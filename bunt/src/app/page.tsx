'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { createClient } from '@/lib/supabase'

const isDemoMode = !process.env.NEXT_PUBLIC_SUPABASE_URL?.startsWith('http') ||
  process.env.NEXT_PUBLIC_SUPABASE_URL?.startsWith('your-')

export default function LoginPage() {
  const router = useRouter()
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    setLoading(true)

    // 데모 모드: 아무 값이나 입력하면 갤러리로 이동
    if (isDemoMode) {
      router.push('/gallery')
      return
    }

    const supabase = createClient()
    const { error } = await supabase.auth.signInWithPassword({ email, password })

    if (error) {
      setError('아이디 또는 비밀번호가 올바르지 않습니다.')
      setLoading(false)
      return
    }

    router.push('/gallery')
    router.refresh()
  }

  return (
    <main style={{
      minHeight: '100vh',
      background: '#0a0a0a',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      fontFamily: 'Georgia, serif',
    }}>
      <div style={{ width: '100%', maxWidth: 360, padding: '0 24px' }}>
        {/* 로고 */}
        <div style={{ textAlign: 'center', marginBottom: 40 }}>
          <h1 style={{
            fontSize: 32,
            letterSpacing: 8,
            color: '#d4a853',
            textTransform: 'uppercase',
            margin: 0,
          }}>BUNT</h1>
          <p style={{
            fontSize: 11,
            letterSpacing: 3,
            color: '#888',
            textTransform: 'uppercase',
            marginTop: 6,
          }}>Kunstschule Magok · 마곡원</p>
        </div>

        {/* 폼 */}
        <form onSubmit={handleSubmit}>
          <input
            type="email"
            placeholder="아이디 (이메일)"
            value={email}
            onChange={e => setEmail(e.target.value)}
            required
            style={{
              width: '100%',
              background: '#0a0a0a',
              border: 'none',
              borderBottom: '1px solid #333',
              padding: '12px 0',
              fontSize: 13,
              color: '#e0d4c0',
              fontFamily: 'Georgia, serif',
              outline: 'none',
              marginBottom: 20,
              boxSizing: 'border-box',
            }}
          />
          <input
            type="password"
            placeholder="비밀번호"
            value={password}
            onChange={e => setPassword(e.target.value)}
            required
            style={{
              width: '100%',
              background: '#0a0a0a',
              border: 'none',
              borderBottom: '1px solid #333',
              padding: '12px 0',
              fontSize: 13,
              color: '#e0d4c0',
              fontFamily: 'Georgia, serif',
              outline: 'none',
              marginBottom: 28,
              boxSizing: 'border-box',
            }}
          />

          {error && (
            <p style={{
              fontSize: 11,
              color: '#c0392b',
              letterSpacing: 1,
              marginBottom: 16,
              textAlign: 'center',
            }}>{error}</p>
          )}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%',
              background: loading ? '#8a6c35' : '#d4a853',
              color: '#0a0a0a',
              border: 'none',
              padding: '14px',
              fontSize: 11,
              letterSpacing: 4,
              textTransform: 'uppercase',
              fontFamily: 'Georgia, serif',
              cursor: loading ? 'not-allowed' : 'pointer',
            }}
          >
            {loading ? '로그인 중...' : '로그인 — Login'}
          </button>
        </form>

        {isDemoMode ? (
          <div style={{ textAlign: 'center', marginTop: 24 }}>
            <p style={{ fontSize: 10, color: '#d4a85360', letterSpacing: 1, marginBottom: 16 }}>
              데모 모드 · 아무 값이나 입력 후 로그인
            </p>
            <button
              onClick={() => router.push('/admin')}
              style={{
                background: 'none', border: '1px solid #2a2a2a',
                color: '#555', fontSize: 9, letterSpacing: 3,
                textTransform: 'uppercase', padding: '7px 16px',
                cursor: 'pointer', fontFamily: 'Georgia, serif',
              }}
            >관리자 페이지 →</button>
          </div>
        ) : (
          <p style={{
            textAlign: 'center', marginTop: 24,
            fontSize: 10, color: '#666', letterSpacing: 1,
          }}>계정 문의는 담당 선생님께</p>
        )}
      </div>
    </main>
  )
}
