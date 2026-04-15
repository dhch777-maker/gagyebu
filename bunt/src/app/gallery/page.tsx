import NavBar from '@/components/NavBar'
import GalleryGrid from '@/components/GalleryGrid'
import { Artwork } from '@/types'

const DEMO_ARTWORKS: Artwork[] = [
  { id: '1', student_id: 'demo', title: '봄의 정원', image_path: '', created_at: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString() },
  { id: '2', student_id: 'demo', title: '나의 가족', image_path: '', created_at: '2025-02-14T00:00:00Z' },
  { id: '3', student_id: 'demo', title: '자화상', image_path: '', created_at: '2025-01-20T00:00:00Z' },
  { id: '4', student_id: 'demo', title: '바다', image_path: '', created_at: '2024-12-05T00:00:00Z' },
  { id: '5', student_id: 'demo', title: '상상의 동물', image_path: '', created_at: '2024-11-15T00:00:00Z' },
  { id: '6', student_id: 'demo', title: '겨울 나무', image_path: '', created_at: '2024-10-20T00:00:00Z' },
  { id: '7', student_id: 'demo', title: '우리 동네', image_path: '', created_at: '2024-09-10T00:00:00Z' },
  { id: '8', student_id: 'demo', title: '친구 초상', image_path: '', created_at: '2024-08-22T00:00:00Z' },
  { id: '9', student_id: 'demo', title: '추상 감정', image_path: '', created_at: '2024-07-15T00:00:00Z' },
]

export default async function GalleryPage() {
  let artworks: Artwork[] = DEMO_ARTWORKS
  let studentName = '최동해'

  try {
    if (
      process.env.NEXT_PUBLIC_SUPABASE_URL &&
      !process.env.NEXT_PUBLIC_SUPABASE_URL.startsWith('your-')
    ) {
      const { createClient } = await import('@/lib/supabase-server')
      const supabase = await createClient()
      const { data: { user } } = await supabase.auth.getUser()

      if (user) {
        const { data: parentAccount } = await supabase
          .from('parent_accounts')
          .select('student_id, display_name, students(name)')
          .eq('id', user.id)
          .single()

        if (parentAccount) {
          studentName = (parentAccount.students as unknown as { name: string } | null)?.name ?? studentName
          const { data: artworkData } = await supabase
            .from('artworks')
            .select('*')
            .eq('student_id', parentAccount.student_id)
            .order('created_at', { ascending: false })

          if (artworkData && artworkData.length > 0) artworks = artworkData
        }
      }
    }
  } catch {
    // Supabase 미연결 시 데모 데이터 사용
  }

  return (
    <div style={{ minHeight: '100vh', background: '#0a0a0a', fontFamily: 'Georgia, serif' }}>
      <NavBar />
      <main style={{ maxWidth: 900, margin: '0 auto', padding: '32px 24px' }}>
        <div style={{ marginBottom: 28, paddingBottom: 16, borderBottom: '1px solid #1e1e1e' }}>
          <p style={{ fontSize: 10, letterSpacing: 3, color: '#888', textTransform: 'uppercase', marginBottom: 6 }}>
            Welcome · 학부모님
          </p>
          <h1 style={{ fontSize: 24, color: '#e0d4c0', letterSpacing: 2 }}>
            <span style={{ color: '#d4a853' }}>{studentName}</span> 작품 갤러리
          </h1>
          <p style={{ fontSize: 10, color: '#888', letterSpacing: 1, marginTop: 4 }}>
            총 {artworks.length}점
          </p>
        </div>
        <GalleryGrid artworks={artworks} />
      </main>
    </div>
  )
}
